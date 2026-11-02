const { app, BrowserWindow, Menu, ipcMain, nativeTheme, shell, session } = require("electron");
const path = require("node:path");
const windowStateKeeper = require("electron-window-state");
const { classifyWindowOpen, classifyNavigation } = require("./linkPolicy");
const { mailboxOrigin, startUrl, isSignedOutPage } = require("./startUrl");
const configOptions = require("../config/options");
const AppTray = require("../tray");
const UnreadService = require("../unread");
const NotificationService = require("../notifications");
const ConnectionMonitor = require("../connection");
const DownloadFeedback = require("../downloads");
const {
  isIntegratedTitleBar,
  titleBarWindowOptions,
  sanitizeOverlayUpdate,
  overlayColorsFromPixel,
  sampleRect,
} = require("./titleBar");

const ICON_PATH = path.join(__dirname, "..", "assets", "icons", "icon-512.png");

let window = null;
let tray = null;
// Set when the user really quits (tray menu, Ctrl+Q, session end), so the
// close-to-tray handler lets the window close.
let quitting = false;

function openExternal(url) {
  // A scheme with no handler rejects; uncaught, that would reach the
  // process-wide unhandledRejection handler. Log the scheme only: the OS
  // error message can echo the URL.
  shell.openExternal(url).catch(() => {
    let scheme = "unparseable";
    try {
      scheme = new URL(url).protocol;
    } catch {
      // keep "unparseable"
    }
    console.error("[LINK] Could not open link externally", { scheme });
  });
}

// Pop-outs (compose, reading pane) get no preload and keep the sandbox.
function webPreferences(config) {
  return {
    partition: config.partition,
    spellcheck: true,
    contextIsolation: true,
    nodeIntegration: false,
    sandbox: true,
    webviewTag: false,
    // Outlook only fires calendar reminders and new-mail notifications while
    // its timers run, so it must keep running at full speed when hidden.
    backgroundThrottling: false,
  };
}

// The main window adds the preload, bundled into one file because a sandboxed
// preload can only require "electron" (see scripts/buildPreload.js). The
// sandbox stays on: Chromium may start Outlook in a renderer process it
// created for Outlook's service worker, which is sandboxed whatever the
// window asks for, so the preload has to work sandboxed anyway.
function mainWebPreferences(config) {
  return {
    ...webPreferences(config),
    preload: path.join(__dirname, "..", "browser", "preload.bundle.js"),
  };
}

function isMainWindowSender(event) {
  return window && !window.isDestroyed() && event.sender === window.webContents;
}

function registerIpcHandlers(config) {
  // The subset of the startup config the main window's preload reads (synchronous: the preload needs it before the page's scripts run)
  ipcMain.on("get-renderer-config", (event) => {
    event.returnValue = isMainWindowSender(event)
      ? { titleBar: config.titleBar, notifications: config.notifications, trayIcon: config.trayIcon }
      : {};
  });


  if (!isIntegratedTitleBar(config)) return;
  // Integrated title bar: match the window controls to the Outlook top bar's height and colour
  ipcMain.on("title-bar-overlay-update", async (event, payload) => {
    if (!isMainWindowSender(event)) return;
    const update = sanitizeOverlayUpdate(payload);
    if (!update) return;
    const { sample, ...overlay } = update;
    if (sample) {
      Object.assign(overlay, await samplePageColors(update.height));
    }
    try {
      if (!window.isDestroyed()) window.setTitleBarOverlay(overlay);
    } catch (err) {
      console.warn("[TITLE_BAR] Could not update the window controls", { error: err.message });
    }
  });
}

// Averages the page behind the window controls down to one pixel. Returns {}
// (keep the renderer's fallback colours) when there is nothing to sample yet,
// e.g. before the window is first shown, when capturePage fails.
async function samplePageColors(height) {
  try {
    const [contentWidth] = window.getContentSize();
    const image = await window.webContents.capturePage(sampleRect(contentWidth, height));
    if (image.isEmpty()) return {};
    const pixel = image.resize({ width: 1, height: 1, quality: "best" }).toBitmap();
    return overlayColorsFromPixel(pixel) ?? {};
  } catch {
    return {};
  }
}

// Applied to the main window and to every pop-out it opens, so links clicked
// inside a pop-out follow the same rules.
function applyLinkPolicy(webContents, config) {
  webContents.setWindowOpenHandler(({ url }) => {
    const action = classifyWindowOpen(url);
    if (action === "child") {
      return {
        action: "allow",
        overrideBrowserWindowOptions: {
          parent: window,
          autoHideMenuBar: true,
          icon: ICON_PATH,
          webPreferences: webPreferences(config),
        },
      };
    }
    if (action === "external") openExternal(url);
    return { action: "deny" };
  });

  webContents.on("did-create-window", (child) => {
    child.webContents.setUserAgent(config.chromeUserAgent);
    applyLinkPolicy(child.webContents, config);
  });

  // See classifyNavigation for why main-frame https navigation is allowed.
  webContents.on("will-navigate", (event, url) => {
    const action = classifyNavigation(url);
    if (action === "allow") return;
    event.preventDefault();
    if (action === "external") openExternal(url);
  });
}

function configureSpellChecker(config) {
  const languages = config.spellCheckerLanguages;
  if (!Array.isArray(languages) || languages.length === 0) return;
  const ses = session.fromPartition(config.partition);
  const available = new Set(ses.availableSpellCheckerLanguages);
  const valid = languages.filter((lang) => available.has(lang));
  if (valid.length !== languages.length) {
    console.warn("[SPELLCHECK] Ignoring unsupported languages", {
      ignored: languages.length - valid.length,
    });
  }
  ses.setSpellCheckerLanguages(valid);
}

const MAILBOX_ORIGIN_KEY = "lastMailboxOrigin";

// A configured url that itself lands on a signed-out page would otherwise
// reload forever.
const SIGN_OUT_RELOAD_INTERVAL_MS = 10_000;
const MAILBOX_SETTLE_MS = 5_000;

// Remembers which Outlook the user ends up in (personal or work) so the next
// launch opens it directly; see app/mainAppWindow/startUrl.js. On a signed-out
// dead end (after signing out, or a remembered mailbox whose session expired)
// it forgets the mailbox and goes back to sign-in, so the next sign-in can be
// either account type.
function trackMailbox(webContents, config, settingsStore) {
  const remember = () => {
    if (webContents.isDestroyed()) return;
    const origin = mailboxOrigin(webContents.getURL());
    if (origin && settingsStore.get(MAILBOX_ORIGIN_KEY) !== origin) {
      settingsStore.set(MAILBOX_ORIGIN_KEY, origin);
    }
  };
  // Landing on a mailbox URL proves nothing: signed out, .../mail/ loads and
  // then redirects to sign-in or to the marketing site. Only a page still on
  // the mailbox a while after loading, or Outlook routing inside itself
  // (which only the signed-in app does), counts.
  let settleTimer = null;
  webContents.on("did-finish-load", () => {
    clearTimeout(settleTimer);
    settleTimer = setTimeout(remember, MAILBOX_SETTLE_MS);
  });
  webContents.on("did-navigate-in-page", (_event, _url, isMainFrame) => {
    if (isMainFrame) remember();
  });

  // did-finish-load rather than did-navigate: the sign-out page signs the
  // session out of other Microsoft services from frames of its own, and
  // leaving early could cut that short.
  let lastReload = 0;
  webContents.on("did-finish-load", () => {
    if (!isSignedOutPage(webContents.getURL())) return;
    settingsStore.delete(MAILBOX_ORIGIN_KEY);
    if (Date.now() - lastReload < SIGN_OUT_RELOAD_INTERVAL_MS) {
      console.warn("[AUTH] Signed-out page reached again right after reloading; staying put");
      return;
    }
    lastReload = Date.now();
    console.info("[AUTH] Signed out, loading sign-in again");
    webContents.loadURL(config.url, { userAgent: config.chromeUserAgent }).catch(() => {});
  });
}

function createWindow(config, settingsStore) {
  const windowState = windowStateKeeper({
    defaultWidth: 1280,
    defaultHeight: 860,
  });

  window = new BrowserWindow({
    title: "Outlook for Linux",
    x: windowState.x,
    y: windowState.y,
    width: windowState.width,
    height: windowState.height,
    minWidth: 640,
    minHeight: 480,
    show: false,
    icon: ICON_PATH,
    autoHideMenuBar: config.menubarAutoHide,
    ...titleBarWindowOptions(config, nativeTheme.shouldUseDarkColors),
    webPreferences: mainWebPreferences(config),
  });
  windowState.manage(window);
  registerIpcHandlers(config);

  let unread = null;
  if (config.trayIcon) {
    tray = new AppTray(window, ICON_PATH, quit);
    unread = new UnreadService(() => window, tray);
    unread.attach(session.fromPartition(config.partition));
  }
  if (config.notifications) {
    new NotificationService(() => window, ICON_PATH).initialize();
  }

  window.once("ready-to-show", () => {
    // With a tray, starting minimized means starting hidden in it.
    if (!config.minimized) {
      window.show();
    } else if (!tray) {
      window.minimize();
    }
  });

  // Closing hides to the tray so Outlook keeps running for notifications and
  // reminders; Quit in the tray menu (or quitting the app) closes for real.
  window.on("close", (event) => {
    if (quitting || !tray || !config.closeToTray) return;
    event.preventDefault();
    window.hide();
  });

  new ConnectionMonitor(window.webContents, config.chromeUserAgent).initialize();
  new DownloadFeedback(window, ICON_PATH).attach(session.fromPartition(config.partition));

  // A new page (reload, sign-out) has no unread count until Outlook reports one.
  window.webContents.on("did-navigate", () => unread?.reset());

  configureSpellChecker(config);
  applyLinkPolicy(window.webContents, config);
  const url = startUrl(config.url, configOptions.url.default, settingsStore.get(MAILBOX_ORIGIN_KEY));
  trackMailbox(window.webContents, config, settingsStore);
  window.loadURL(url, { userAgent: config.chromeUserAgent });
  return window;
}

function buildMenu() {
  const template = [
    { role: "fileMenu" },
    { role: "editMenu" },
    { role: "viewMenu" },
    { role: "windowMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function quit() {
  quitting = true;
  app.quit();
}

app.on("before-quit", () => {
  quitting = true;
});

function onAppReady(config, settingsStore) {
  buildMenu();
  return createWindow(config, settingsStore);
}

function show() {
  if (!window || window.isDestroyed()) return;
  if (window.isMinimized()) window.restore();
  window.show();
  window.focus();
}

module.exports = {
  onAppReady,
  show,
  get window() {
    return window;
  },
};
