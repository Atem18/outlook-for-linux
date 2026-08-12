'use strict';

const {
  app,
  BrowserWindow,
  Tray,
  Menu,
  ipcMain,
  shell,
  session,
  Notification,
} = require('electron');
const path = require('path');
const { generateTrayIcon, ICON_PATH } = require('./icon-generator');

// ─── Desktop file association (before app.whenReady) ─────────────────────────
// Tells Wayland/KDE which .desktop file owns this process.
// Do NOT call app.setName() — it changes WM_CLASS and breaks StartupWMClass matching.
if (typeof app.setDesktopFileName === 'function') {
  app.setDesktopFileName('outlook-for-linux');       // Electron 28+
} else {
  process.env.CHROME_DESKTOP = 'outlook-for-linux.desktop'; // older Electron
}

// ─── Wayland / Ozone ──────────────────────────────────────────────────────────
if (!process.env.ELECTRON_OZONE_PLATFORM_HINT) {
  app.commandLine.appendSwitch('ozone-platform-hint', 'auto');
}
app.commandLine.appendSwitch('enable-features', 'WaylandWindowDecorations');
app.commandLine.appendSwitch('enable-wayland-ime');

// ─── Single instance ──────────────────────────────────────────────────────────
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) { app.quit(); process.exit(0); }

// ─── State ────────────────────────────────────────────────────────────────────
let mainWindow  = null;
let tray        = null;
let unreadCount = -1;
let isQuitting  = false;
let pollTimer   = null;

// Microsoft uses both domains — office.com is classic, cloud.microsoft is new Outlook
const OUTLOOK_ORIGINS = [
  'https://outlook.office.com',
  'https://outlook.cloud.microsoft',
  'https://outlook.office365.com',
];
const OUTLOOK_URL = 'https://outlook.office.com/mail/';
function isOutlookPage(url) {
  return OUTLOOK_ORIGINS.some((o) => (url || '').startsWith(o));
}
const USER_AGENT =
  `Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) ` +
  `Chrome/${process.versions.chrome} Safari/537.36`;

// ─── DOM scraper ──────────────────────────────────────────────────────────────
// Runs per-frame via WebFrameMain.executeJavaScript() — the main process calls
// it on every frame in the tree so it can reach cross-origin child frames
// (e.g. the Microsoft 365 suite shell loads Outlook inside such a frame).
// Counting [aria-label^="Unread "] items is kept as a last resort only —
// Outlook virtualises the email list so it returns a lower bound, not the total.

const DOM_SCRAPER = `(function () {
  // No \\b after the inbox name — new Outlook concatenates "Inbox83unread" with
  // no whitespace, so a word-boundary before the digit never fires.
  const INBOX_RE      = /(inbox|boîte de réception|posteingang|bandeja de entrada|caixa de entrada|posta in arrivo|postvak in|inkorgen|innboks|indbakke|saapuneet|skrzynka odbiorcza)/i;
  const INBOX_ARIA_RE = /^(inbox|boîte de réception|posteingang|bandeja de entrada)/i;
  const UNREAD_KW     = /unread|non lus?|ungelesen|sin leer/i;

  function $$(sel, root) {
    const res = [...root.querySelectorAll(sel)];
    for (const el of root.querySelectorAll('*')) {
      if (el.shadowRoot) res.push(...$$(sel, el.shadowRoot));
    }
    return res;
  }

  // ── Strategy 1: title "(N) …" prefix ────────────────────────────────────────
  const tm = document.title.match(/^\\((\\d+)\\)/);
  if (tm) return { count: parseInt(tm[1], 10), via: 'title' };

  // ── Strategy 2: treeitem whose text includes an inbox name ──────────────────
  for (const el of $$('[role="treeitem"]', document)) {
    const raw = (el.textContent || '').replace(/\\s+/g, ' ').trim();
    if (!INBOX_RE.test(raw)) continue;
    const m = raw.replace(INBOX_RE, '').match(/(\\d+)/);
    return { count: m ? parseInt(m[1], 10) : 0, via: 'treeitem-text' };
  }

  // ── Strategy 3: aria-label on treeitems (classic office.com) ────────────────
  for (const el of $$('[role="treeitem"]', document)) {
    const label = el.getAttribute('aria-label') || '';
    if (!INBOX_ARIA_RE.test(label)) continue;
    if (UNREAD_KW.test(label) || /\\(\\d+\\)/.test(label)) {
      const m = label.match(/(\\d+)/);
      return { count: m ? parseInt(m[1], 10) : 0, via: 'treeitem-aria' };
    }
    return { count: 0, via: 'treeitem-aria-zero' };
  }

  // ── Strategy 4: badge <span>N</span> whose parent text is "NNNunread" ───────
  // New Outlook renders a leaf span with just the count next to a sibling
  // "unread" span; their shared parent collapses to "83unread" with no space.
  for (const span of document.querySelectorAll('span')) {
    if (span.children.length > 0) continue;
    const n = parseInt(span.textContent.trim(), 10);
    if (!(n > 0 && n < 10000)) continue;
    const pText = (span.parentElement?.textContent || '').replace(/\\s+/g, '').toLowerCase();
    if (!/^\\d+unread$/.test(pText) && !/^unread\\d+$/.test(pText)) continue;
    // Walk up to the nearest treeitem and confirm it is an inbox folder
    let node = span.parentElement;
    while (node && node.getAttribute?.('role') !== 'treeitem') node = node.parentElement;
    if (node && INBOX_RE.test(node.textContent)) return { count: n, via: 'inbox-badge' };
  }

  // ── Last resort: count visible unread items (virtual scroll — lower bound) ───
  const els = $$('[aria-label^="Unread "]', document);
  if (els.length > 0) return { count: els.length, via: 'unread-items' };

  return { count: 0, via: 'none' };
})()`

// Collect the WebFrameMain tree depth-first
function collectAllFrames(frame) {
  const result = [frame];
  for (const child of (frame.frames || [])) result.push(...collectAllFrames(child));
  return result;
}

// ─── Boot ─────────────────────────────────────────────────────────────────────
app.whenReady().then(() => {
  createTray();
  createWindow();
  setupIPC();
  setupSessionHooks();
});

app.on('second-instance', showWindow);
app.on('before-quit',       () => { isQuitting = true; });
app.on('window-all-closed', (e) => { if (!isQuitting) e.preventDefault(); });

// ─── Main window ──────────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width:           1280,
    height:          900,
    minWidth:        800,
    minHeight:       600,
    title:           'Outlook for Linux',
    icon:            ICON_PATH,
    autoHideMenuBar: true,
    backgroundColor: '#0078d4',
    frame:           false,
    titleBarStyle:   'hidden',
    titleBarOverlay: {
      color:       '#efefef',
      symbolColor: '#323130',
      height:      48,
    },
    webPreferences: {
      preload:          path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration:  false,
      sandbox:          true,
      partition:        'persist:outlook',
    },
  });

  mainWindow.webContents.setUserAgent(USER_AGENT);
  mainWindow.loadURL(OUTLOOK_URL);

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (isOutlookPage(url)) return { action: 'allow' };
    // Outlook popups (Reminders, etc.) open as about:blank then navigate internally
    if (url === 'about:blank') return { action: 'allow' };
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.webContents.on('did-create-window', (childWindow) => {
    childWindow.webContents.setUserAgent(USER_AGENT);

    childWindow.webContents.setWindowOpenHandler(({ url }) => {
      if (isOutlookPage(url) || url === 'about:blank') return { action: 'allow' };
      shell.openExternal(url);
      return { action: 'deny' };
    });

    // Allow child windows to navigate to Outlook and auth pages; send anything
    // else to the default browser rather than opening it inside Electron.
    childWindow.webContents.on('will-navigate', (_e, url) => {
      if (url.startsWith('about:')) return;
      const allowed = [
        'outlook.office.com', 'outlook.office365.com',
        'outlook.cloud.microsoft',
        'login.microsoftonline.com', 'login.microsoft.com',
        'login.live.com', 'account.microsoft.com',
      ];
      try {
        const { hostname } = new URL(url);
        if (!allowed.some((d) => hostname.endsWith(d))) {
          _e.preventDefault();
          shell.openExternal(url);
        }
      } catch (e) { console.warn('[nav] invalid child URL', url, e.message); }
    });
  });

  mainWindow.webContents.on('will-navigate', (_e, url) => {
    // about: URLs are internal browser pages, never open externally
    if (url.startsWith('about:')) return;
    const allowed = [
      'outlook.office.com', 'outlook.office365.com',
      'outlook.cloud.microsoft',          // new Outlook domain
      'login.microsoftonline.com', 'login.microsoft.com',
      'login.live.com', 'account.microsoft.com',
    ];
    try {
      const { hostname } = new URL(url);
      if (!allowed.some((d) => hostname.endsWith(d))) {
        _e.preventDefault();
        shell.openExternal(url);
      }
    } catch (e) { console.warn('[nav] invalid URL', url, e.message); }
  });

  // Only start polling once we're actually on an Outlook page —
  // avoids the poll timer being reset on every Microsoft SSO redirect.
  mainWindow.webContents.on('did-finish-load', () => {
    const url = mainWindow.webContents.getURL();
    if (isOutlookPage(url)) {
      startPolling();
      // Electron's titleBarOverlay draws the min/max/close buttons but doesn't
      // report `display-mode: window-controls-overlay`, so Outlook's own WCO-aware
      // CSS never activates and the window has no drag region. #OwaTitleBar is
      // Outlook's own (stable, purpose-built) title bar row — make it draggable
      // directly, excluding anything interactive inside it.
      mainWindow.webContents.insertCSS(`
        #OwaTitleBar { -webkit-app-region: drag; }
        #OwaTitleBar button,
        #OwaTitleBar a,
        #OwaTitleBar input,
        #OwaTitleBar [role="button"],
        #OwaTitleBar [role="menuitem"],
        #OwaTitleBar [tabindex] { -webkit-app-region: no-drag; }
      `).catch((e) => console.warn('[titlebar] insertCSS failed', e.message));
    } else {
      // Still on login flow — just check the title in case it carries a count
      pollDOM();
    }
  });

  // Re-poll on SPA navigations within Outlook
  mainWindow.webContents.on('did-navigate-in-page', () => {
    if (isOutlookPage(mainWindow.webContents.getURL())) {
      pollDOM();
    }
  });

  // Catch title changes (preload fast-path AND backup for when window is shown)
  mainWindow.webContents.on('page-title-updated', (_e, title) => {
    const m = (title || '').match(/^\((\d+)\)/);
    if (m) applyUnreadCount(parseInt(m[1], 10));
  });

  mainWindow.on('close',  (e) => { if (!isQuitting) { e.preventDefault(); mainWindow.hide(); } });
  mainWindow.on('closed', ()  => { mainWindow = null; stopPolling(); });
}

// ─── DOM polling ──────────────────────────────────────────────────────────────
let diagRun = 0; // run diagnostics a few times to catch lazy-loaded nav

async function pollDOM() {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  try {
    const frames = collectAllFrames(mainWindow.webContents.mainFrame);
    let best = { count: 0, via: 'none' };

    for (const frame of frames) {
      let result;
      try {
        result = await frame.executeJavaScript(DOM_SCRAPER);
      } catch (e) {
        continue; // frame loading / navigating / destroyed
      }
      if (!result) continue;
      // Prefer any reliable strategy over the visible-items last resort
      if (result.via !== 'none' && result.via !== 'unread-items') { best = result; break; }
      if (best.via === 'none' && result.via === 'unread-items')     best = result;
    }

    const n = Math.max(0, best.count);
    console.log(`[poll] count=${n} via=${best.via} title="${mainWindow.webContents.getTitle()}"`);
    applyUnreadCount(n);

    // When all reliable strategies fail, dump DOM info (up to 5 times) to help find the right selector
    if (diagRun < 5 && (best.via === 'unread-items' || best.via === 'none')) {
      diagRun++;
      const DIAG = `(function () {
        const total = document.querySelectorAll('*').length;

        // All unique aria-labels on the page — unfiltered
        const allAria = [...new Set(
          [...document.querySelectorAll('[aria-label]')]
            .map(el => el.getAttribute('aria-label').trim())
            .filter(v => v.length > 0 && v.length < 150)
        )].slice(0, 50);

        // All roles in use
        const allRoles = [...new Set(
          [...document.querySelectorAll('[role]')].map(el => el.getAttribute('role'))
        )];

        // Leaf elements whose text is just a short number (potential badge)
        const numericLeaves = [...document.querySelectorAll('*')]
          .filter(el => el.children.length === 0 && /^\\d+$/.test(el.textContent.trim())
                        && parseInt(el.textContent) > 0 && parseInt(el.textContent) < 10000)
          .slice(0, 10)
          .map(el => ({
            tag: el.tagName, n: el.textContent.trim(),
            pAria: el.parentElement && el.parentElement.getAttribute('aria-label'),
            pRole: el.parentElement && el.parentElement.getAttribute('role'),
            pText: el.parentElement && el.parentElement.textContent.replace(/\\s+/g, ' ').trim().slice(0, 80)
          }));

        return { url: location.href, title: document.title, total, allRoles, numericLeaves, allAria };
      })()`;
      console.log('[diag] total frames:', frames.length);
      for (const frame of frames) {
        try {
          const d = await frame.executeJavaScript(DIAG);
          if (d) console.log('[diag] frame ok', JSON.stringify(d, null, 2));
        } catch (e) {
          console.log('[diag] frame err', frame.url, e.message);
        }
      }
    }
  } catch (e) {
    console.log(`[poll] error: ${e.message}`);
  }
}

function startPolling() {
  // Don't reset the interval if it's already running on an Outlook page
  if (pollTimer) return;
  pollDOM();                                        // immediate first read
  pollTimer = setInterval(pollDOM, 5_000);
}

function stopPolling() {
  if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
}

// ─── Apply count ──────────────────────────────────────────────────────────────
function applyUnreadCount(count) {
  if (count === unreadCount) return;
  const previous = unreadCount;
  unreadCount    = count;
  const n        = Math.max(0, count);

  app.setBadgeCount(n);
  console.log(`Badge count set to '${n}'`);

  updateTrayIcon(n);

  if (previous >= 0 && count > previous && mainWindow) {
    // Demand attention on the taskbar button (KDE "glow" effect)
    if (!mainWindow.isFocused()) mainWindow.flashFrame(true);

    if (!mainWindow.isVisible() && Notification.isSupported()) {
      const diff  = count - previous;
      const notif = new Notification({
        title:   'Outlook for Linux – New mail',
        body:    `${diff} new message${diff !== 1 ? 's' : ''} (${count} unread total)`,
        icon:    ICON_PATH,
        urgency: 'normal',
      });
      notif.on('click', showWindow);
      notif.show();
    }
  }
}

// ─── Tray ─────────────────────────────────────────────────────────────────────
function createTray() {
  tray = new Tray(generateTrayIcon(0));
  tray.setToolTip('Outlook for Linux');
  tray.setContextMenu(buildMenu());
  tray.on('click',        toggleWindow);
  tray.on('double-click', toggleWindow);
}

function buildMenu() {
  const n     = Math.max(0, unreadCount);
  const label = n > 0 ? `${n} unread message${n !== 1 ? 's' : ''}` : 'No unread messages';
  return Menu.buildFromTemplate([
    { label, enabled: false },
    { type: 'separator' },
    { label: 'Open Outlook', click: showWindow },
    {
      label: 'New Email',
      click: () => {
        showWindow();
        mainWindow?.webContents.loadURL('https://outlook.office.com/mail/deeplink/compose');
      },
    },
    { type: 'separator' },
    { label: 'Reload',  click: () => mainWindow?.webContents.reload() },
    { type: 'separator' },
    { label: 'Quit',    click: () => { isQuitting = true; app.quit(); } },
  ]);
}

function updateTrayIcon(count) {
  if (!tray) return;
  tray.setImage(generateTrayIcon(count));
  tray.setToolTip(count > 0
    ? `Outlook for Linux – ${count} unread message${count !== 1 ? 's' : ''}`
    : 'Outlook for Linux – No unread messages'
  );
  tray.setContextMenu(buildMenu());
}

// ─── Window helpers ───────────────────────────────────────────────────────────
function showWindow() {
  if (!mainWindow) createWindow();
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.show();
  mainWindow.focus();
  mainWindow.flashFrame(false);
}

function toggleWindow() {
  if (!mainWindow) return showWindow();
  mainWindow.isVisible() && mainWindow.isFocused() ? mainWindow.hide() : showWindow();
}

// ─── IPC ─────────────────────────────────────────────────────────────────────
function setupIPC() {
  ipcMain.on('tray-update', (_e, { count }) => {
    applyUnreadCount(Math.max(0, parseInt(count, 10) || 0));
  });
}

// ─── Session hooks ────────────────────────────────────────────────────────────
function setupSessionHooks() {
  const ses = session.fromPartition('persist:outlook');
  ses.webRequest.onBeforeRequest(
    { urls: [
      '*://browser.pipe.aria.microsoft.com/*',
      '*://*.clarity.ms/*',
      '*://dc.services.visualstudio.com/*',
    ]},
    (_d, cb) => cb({ cancel: true })
  );
}
