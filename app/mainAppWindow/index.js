const { app, BrowserWindow, Menu, ipcMain, nativeTheme, shell, session } = require("electron");
const path = require("node:path");
const fs = require("node:fs");
const { execFile } = require("node:child_process");
const windowStateKeeper = require("electron-window-state");
const { classifyWindowOpen, classifyNavigation } = require("./linkPolicy");
const { mailboxOrigin, startUrl, isSignedOutPage } = require("./startUrl");
const { PROTOCOL_SCHEMES, parseMailto, findMailtoArg, composeUrl } = require("./mailtoLink");
const { claimFileTypes } = require("./fileTypeDefault");
const { resolveExternalUrl } = require("./teamsLink");
const configOptions = require("../config/options");
const AppTray = require("../tray");
const { UnreadService, UnreadTotal } = require("../unread");
const NotificationService = require("../notifications");
const ConnectionMonitor = require("../connection");
const DownloadFeedback = require("../downloads");
const Zoom = require("../zoom");
const CustomCSS = require("../customCSS");
const AppIcon = require("../appIcon");
const { AccountStore } = require("../accounts/store");
const AccountViews = require("../accounts/views");
const EmlViewer = require("../emlViewer");
const { findEmlArgs } = require("../emlViewer/message");
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
let zoom = null;
let customCSS = null;
let appIcon = null;
let appConfig = null;
let store = null;
// Multiple accounts mode (config.multiAccount): the account list and the
// views of the accounts other than the main window's own page.
let accounts = null;
let accountViews = null;
let unreadTotal = null;
let notifications = null;
let downloads = null;
// A mailto: link that arrived before the main window existed.
let pendingMailto = null;
// .eml files that arrived before the app was ready, and their viewer.
const pendingEmlFiles = [];
let emlViewer = null;
// webContents of open mailto: compose windows. They load the preload for
// their title bar only.
const composeContents = new Set();
// The main window's header colour, sampled from its top rows; compose windows
// paint their title bar strip with it.
let mainHeaderColor = null;
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

function runCommand(command, args) {
  return new Promise((resolve, reject) =>
    execFile(command, args, { timeout: 10_000 }, (err, stdout) => (err ? reject(err) : resolve(stdout))),
  );
}

function isSandboxed() {
  return Boolean(process.env.FLATPAK_ID || process.env.SNAP_INSTANCE_NAME);
}

// A link leaving the app from Outlook's windows. Teams meeting links go to
// the Teams desktop app when the teamsLinks option says so (./teamsLink.js).
function openOutboundLink(url, config) {
  resolveExternalUrl({ url, mode: config.teamsLinks, sandboxed: isSandboxed(), run: runCommand })
    .then((target) => {
      if (target !== url) console.info("[LINK] Opening Teams link in the Teams app");
      openExternal(target);
    })
    .catch(() => openExternal(url));
}

// Pop-outs (compose, reading pane) get no preload and keep the sandbox. They
// share the session of the account that opened them.
function webPreferences(config, partition = config.partition) {
  return {
    partition,
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
function mainWebPreferences(config, partition = config.partition) {
  return {
    ...webPreferences(config, partition),
    preload: path.join(__dirname, "..", "browser", "preload.bundle.js"),
  };
}

// The visible account's Outlook page: the window's own page, or with
// multiple accounts the view of the account on screen.
function activeContents() {
  if (accountViews) return accountViews.activeContents;
  return window && !window.isDestroyed() ? window.webContents : null;
}

// An account's Outlook page: the window's own page or an account view (not
// a pop-out or compose window).
function isOutlookPage(contents) {
  if (!window || window.isDestroyed()) return false;
  if (contents === window.webContents) return true;
  return Boolean(accountViews?.accountOf(contents));
}

function isMainWindowSender(event) {
  return isOutlookPage(event.sender);
}

// The page a menu item acts on: the visible account's in the main window,
// the window's own page in a pop-out.
function contentsForWindow(win) {
  if (!win || win === window) return activeContents();
  return win.isDestroyed() ? null : win.webContents;
}

function isComposeSender(event) {
  return composeContents.has(event.sender);
}

function registerIpcHandlers(config) {
  // The subset of the startup config the main window's preload reads (synchronous: the preload needs it before the page's scripts run)
  ipcMain.on("get-renderer-config", (event) => {
    if (isMainWindowSender(event)) {
      event.returnValue = {
        titleBar: config.titleBar,
        notifications: config.notifications,
        trayIcon: config.trayIcon,
        webauthn: config.webauthn,
        webauthnExtraOrigins: config.webauthnExtraOrigins,
      };
    } else if (isComposeSender(event)) {
      // Notifications and the unread count belong to the main window alone.
      event.returnValue = isIntegratedTitleBar(config)
        ? { titleBar: config.titleBar, popOutTitleBar: { color: mainHeaderColor } }
        : {};
    } else {
      event.returnValue = {};
    }
  });


  if (!isIntegratedTitleBar(config)) return;
  // Integrated title bar: match the window controls to the Outlook top bar's height and colour
  ipcMain.on("title-bar-overlay-update", async (event, payload) => {
    const outlookPage = isMainWindowSender(event);
    if (!outlookPage && !isComposeSender(event)) return;
    const contents = event.sender;
    const target = outlookPage ? window : BrowserWindow.fromWebContents(contents);
    if (!target) return;
    const update = sanitizeOverlayUpdate(payload);
    if (!update) return;
    const { sample, ...overlay } = update;
    if (sample) {
      Object.assign(overlay, await samplePageColors(target, contents, update.height));
    }
    if (outlookPage) {
      // Every account's page reports its bar; only the visible one's applies,
      // and the others' are kept for when they are switched to.
      overlays.set(contents, overlay);
      if (contents !== activeContents()) return;
      refreshMainHeaderColor();
    }
    applyOverlay(target, overlay);
  });
}

// The last window-controls overlay each account's page asked for (weak: a
// removed account's page drops out by itself).
const overlays = new WeakMap();

function applyOverlay(target, overlay) {
  try {
    if (!target.isDestroyed()) target.setTitleBarOverlay(overlay);
  } catch (err) {
    console.warn("[TITLE_BAR] Could not update the window controls", { error: err.message });
  }
}

// The header colour for compose windows' title bar strips. Not the colour
// behind the main window's controls: that end of the bar holds the avatar.
// The top two rows, across the whole width, are bare header (its icons,
// search box and avatar are vertically centred), or the average of a theme
// image. Sampled when a compose window opens (the main window reports its
// controls' colour mostly while still loading) and on those reports; open
// compose windows follow changes, such as a theme switch.
async function refreshMainHeaderColor() {
  const contents = activeContents();
  if (!window || window.isDestroyed() || !contents) return;
  try {
    const [contentWidth] = window.getContentSize();
    const image = await contents.capturePage({ x: 0, y: 1, width: contentWidth, height: 2 });
    if (image.isEmpty()) return;
    const colors = overlayColorsFromPixel(image.resize({ width: 1, height: 1, quality: "best" }).toBitmap());
    const color = colors?.color.slice(0, 7);
    if (!color || color === mainHeaderColor) return;
    mainHeaderColor = color;
    for (const contents of composeContents) contents.send("pop-out-header-color", color);
  } catch {
    // keep the last colour
  }
}

// Averages the page behind the window controls down to one pixel. Returns {}
// (keep the renderer's fallback colours) when there is nothing to sample yet,
// e.g. before the window is first shown, when capturePage fails.
async function samplePageColors(target, contents, height) {
  try {
    const [contentWidth] = target.getContentSize();
    const image = await contents.capturePage(sampleRect(contentWidth, height));
    if (image.isEmpty()) return {};
    const pixel = image.resize({ width: 1, height: 1, quality: "best" }).toBitmap();
    return overlayColorsFromPixel(pixel) ?? {};
  } catch {
    return {};
  }
}

// Applied to the main window and to every pop-out it opens, so links clicked
// inside a pop-out follow the same rules. Pop-outs stay on the session of the
// account that opened them: Outlook writes into the about:blank windows it
// opens, which only works within one session.
function applyLinkPolicy(webContents, config, partition = config.partition) {
  webContents.setWindowOpenHandler(({ url }) => {
    const action = classifyWindowOpen(url);
    if (action === "child") {
      return {
        action: "allow",
        overrideBrowserWindowOptions: {
          parent: window,
          autoHideMenuBar: true,
          icon: appIcon.windowImage(),
          webPreferences: webPreferences(config, partition),
        },
      };
    }
    if (action === "external") openOutboundLink(url, config);
    return { action: "deny" };
  });

  webContents.on("did-create-window", (child) => {
    child.webContents.setUserAgent(config.chromeUserAgent);
    applyLinkPolicy(child.webContents, config, partition);
    attachPageFeatures(child.webContents);
  });

  // See classifyNavigation for why main-frame https navigation is allowed.
  webContents.on("will-navigate", (event, url) => {
    const action = classifyNavigation(url);
    if (action === "allow") return;
    event.preventDefault();
    if (action === "external") openOutboundLink(url, config);
  });
}

// Saved zoom and the user's custom CSS, for every Outlook window.
function attachPageFeatures(webContents) {
  zoom?.attach(webContents);
  customCSS?.attach(webContents);
}

function configureSpellChecker(config, ses) {
  const languages = config.spellCheckerLanguages;
  if (!Array.isArray(languages) || languages.length === 0) return;
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

// Where the main account's remembered mailbox is kept.
function mainMailboxMemory(settingsStore) {
  return {
    get: () => settingsStore.get(MAILBOX_ORIGIN_KEY),
    set: (origin) => settingsStore.set(MAILBOX_ORIGIN_KEY, origin),
    forget: () => settingsStore.delete(MAILBOX_ORIGIN_KEY),
  };
}

// Where an added account's remembered mailbox is kept.
function accountMailboxMemory(id) {
  return {
    get: () => accounts.get(id)?.mailboxOrigin,
    set: (origin) => accounts.get(id) && accounts.setMailboxOrigin(id, origin),
    forget: () => accounts.get(id) && accounts.setMailboxOrigin(id, null),
  };
}

// Remembers which Outlook the user ends up in (personal or work) so the next
// launch opens it directly; see app/mainAppWindow/startUrl.js. On a signed-out
// dead end (after signing out, or a remembered mailbox whose session expired)
// it forgets the mailbox and goes back to sign-in, so the next sign-in can be
// either account type. One per account.
function trackMailbox(webContents, config, memory) {
  const remember = () => {
    if (webContents.isDestroyed()) return;
    const origin = mailboxOrigin(webContents.getURL());
    if (origin && memory.get() !== origin) memory.set(origin);
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
    memory.forget();
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
    icon: appIcon.windowImage(),
    autoHideMenuBar: config.menubarAutoHide,
    ...titleBarWindowOptions(config, nativeTheme.shouldUseDarkColors),
    webPreferences: mainWebPreferences(config),
  });
  windowState.manage(window);
  registerIpcHandlers(config);
  if (multiAccountEnabled(config)) accounts = new AccountStore(settingsStore, config.partition);

  if (config.trayIcon) tray = new AppTray(window, appIcon.path, quit);
  // The dock badge and the tray show every account's unread mail; the
  // switcher shows each account's.
  unreadTotal = new UnreadTotal(
    tray,
    () => (window && !window.isDestroyed() ? window.webContents : null),
    (counts) => accountViews?.setUnread(counts)
  );
  notifications = config.notifications
    ? new NotificationService(() => window, appIcon.path, { isOutlookPage, accountOf: notificationAccount })
    : null;
  notifications?.initialize();

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

  downloads = new DownloadFeedback(window, appIcon.path);
  // Window icons follow in AppIcon itself; the menu's checkbox states too.
  appIcon.onChange((iconPath) => {
    tray?.setBaseIconPath(iconPath);
    notifications?.setIconPath(iconPath);
    downloads.setIconPath(iconPath);
    buildMenu();
  });

  zoom = new Zoom(activeContents, settingsStore);
  customCSS = new CustomCSS(config.customCSSLocation);
  const url = setupOutlookPage(window.webContents, config, {
    id: MAIN_ACCOUNT,
    partition: config.partition,
    memory: mainMailboxMemory(settingsStore),
  });
  window.loadURL(url, { userAgent: config.chromeUserAgent });

  if (accounts) {
    accountViews = new AccountViews(window, accounts, {
      webPreferences: (partition) => mainWebPreferences(config, partition),
      userAgent: config.chromeUserAgent,
      setupPage: (contents, account) =>
        setupOutlookPage(contents, config, {
          id: account.id,
          partition: account.partition,
          memory: accountMailboxMemory(account.id),
        }),
      onRemoved: (account) => unreadTotal.delete(account.id),
      onSwitched: () => {
        // The window controls take the visible account's top bar colour.
        const overlay = overlays.get(activeContents());
        if (overlay) applyOverlay(window, overlay);
        if (isIntegratedTitleBar(config)) refreshMainHeaderColor();
        buildMenu();
      },
    });
    accountViews.initialize();
    console.info("[ACCOUNTS] Multiple accounts on", { count: accounts.list().length });
  }
  buildMenu();
  return window;
}

const MAIN_ACCOUNT = "main";

// Multiple accounts, unless Intune is on: the broker signs in one account per
// machine, and its SSO cookie would go to every account's sign-in (as in
// teams-for-linux, the two are exclusive).
function multiAccountEnabled(config) {
  if (!config.multiAccount) return false;
  if (config.intune) {
    console.warn("[ACCOUNTS] multiAccount is ignored while intune is on");
    return false;
  }
  return true;
}

/**
 * Makes a page an account's Outlook page: user agent, link policy, zoom and
 * custom CSS, remembered mailbox and sign-out handling, reload when the
 * network was down, unread count, and, for an added account's session,
 * downloads, spell checking and security keys.
 * @param {Electron.WebContents} contents
 * @param {object} config
 * @param {{id: string, partition: string, memory: {get: Function, set: Function, forget: Function}}} account
 * @returns {string} the URL to load
 */
function setupOutlookPage(contents, config, { id, partition, memory }) {
  const ses = session.fromPartition(partition);
  contents.setUserAgent(config.chromeUserAgent);
  attachPageFeatures(contents);
  applyLinkPolicy(contents, config, partition);
  trackMailbox(contents, config, memory);
  new ConnectionMonitor(contents, config.chromeUserAgent).initialize();
  downloads.attach(ses);
  configureSpellChecker(config, ses);
  if (id !== MAIN_ACCOUNT && process.platform === "linux" && config.webauthn) {
    // The main window's page is covered by initialize() in app/index.js.
    require("../webauthn").attachWebContents(contents);
  }

  const unread = new UnreadService(
    () => (contents.isDestroyed() ? null : contents),
    (count) => {
      // A removed account's page can still report while it closes.
      if (id === MAIN_ACCOUNT || accounts?.get(id)) unreadTotal.set(id, count);
    }
  );
  unread.attach(ses);
  // A new page (reload, sign-out) has no unread count until Outlook reports one.
  contents.on("did-navigate", () => unread.reset());
  contents.once("destroyed", () => unread.destroy());

  return startUrl(config.url, configOptions.url.default, memory.get());
}

// For a notification: the account it came from, when there are several.
function notificationAccount(contents) {
  const account = accountViews?.accountOf(contents);
  if (!account) return null;
  return {
    name: accounts.list().length > 1 ? account.name : null,
    show: () => accountViews.switch(account.id),
  };
}

function buildMenu() {
  const template = [
    {
      label: "File",
      submenu: [
        {
          label: "Open Message…",
          accelerator: "CmdOrCtrl+O",
          click: (_item, win) =>
            emlViewer?.openDialog(win).catch((err) => console.error("[EML] Open dialog failed", { error: err.message })),
        },
        { type: "separator" },
        { label: "Quit", accelerator: "CmdOrCtrl+Q", click: quit },
      ],
    },
    { role: "editMenu" },
    {
      label: "View",
      submenu: [
        // Not the roles: those act on the window's own page, which is not
        // the visible one when another account is shown.
        { label: "Reload", accelerator: "CmdOrCtrl+R", click: (_item, win) => contentsForWindow(win)?.reload() },
        {
          label: "Force Reload",
          accelerator: "Shift+CmdOrCtrl+R",
          click: (_item, win) => contentsForWindow(win)?.reloadIgnoringCache(),
        },
        {
          label: "Toggle Developer Tools",
          accelerator: "Ctrl+Shift+I",
          click: (_item, win) => contentsForWindow(win)?.toggleDevTools(),
        },
        { type: "separator" },
        // Not the zoom roles: these save the zoom for the next launch. The
        // keys are handled by app/zoom (before the page sees them), so the
        // accelerators here are labels only.
        { label: "Actual Size", accelerator: "CmdOrCtrl+0", registerAccelerator: false, click: (_item, win) => zoom?.step(0, contentsForWindow(win)) },
        { label: "Zoom In", accelerator: "CmdOrCtrl+Plus", registerAccelerator: false, click: (_item, win) => zoom?.step(1, contentsForWindow(win)) },
        { label: "Zoom Out", accelerator: "CmdOrCtrl+-", registerAccelerator: false, click: (_item, win) => zoom?.step(-1, contentsForWindow(win)) },
        { type: "separator" },
        { role: "togglefullscreen" },
        { type: "separator" },
        appIcon.menuItem(() => window),
      ],
    },
    ...(accountViews ? [accountsMenu()] : []),
    { role: "windowMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  // Reachable from the tray too: the integrated title bar has no menu bar.
  tray?.setMenuItems([appIcon.menuItem(() => window), ...(accountViews ? [accountsMenu()] : [])]);
}

// Accounts menu, in the menu bar and the tray. Ctrl+Alt+<n> is handled by
// app/accounts/views.js before the page sees the keys; here it is a label.
function accountsMenu() {
  const active = accounts.getActive();
  const run = (action) => () => {
    show();
    action().catch((err) => console.error("[ACCOUNTS] Account action failed", { error: err.message }));
  };
  return {
    label: "Accounts",
    submenu: [
      ...accounts.list().map((account, index) => ({
        label: account.name.replaceAll("&", "&&"),
        type: "radio",
        checked: account.id === active.id,
        accelerator: `Ctrl+Alt+${index + 1}`,
        registerAccelerator: false,
        click: () => {
          show();
          accountViews.switch(account.id);
        },
      })),
      { type: "separator" },
      { label: "Add Account…", click: run(() => accountViews.add()) },
      { label: "Rename This Account…", click: run(() => accountViews.renameActive()) },
      { label: "Remove This Account…", enabled: !active.main, click: run(() => accountViews.removeActive()) },
    ],
  };
}

function quit() {
  quitting = true;
  app.quit();
}

app.on("before-quit", () => {
  quitting = true;
});

// Opt-in (defaultMailClient): claiming mailto: and .eml files silently would
// take them from the user's current mail client. Inside Flatpak and Snap,
// xdg-settings and xdg-mime cannot change the host's handlers, so the
// desktop's own settings are the way there; the desktop entry already lists
// the app for mailto: and message/rfc822.
function registerMailClient(config) {
  if (!config.defaultMailClient) return;
  if (isSandboxed()) {
    console.info("[MAILTO] Sandboxed install: choose the default mail client in the desktop settings");
    return;
  }
  for (const scheme of PROTOCOL_SCHEMES) {
    if (app.isDefaultProtocolClient(scheme)) continue;
    const ok = app.setAsDefaultProtocolClient(scheme);
    console.info("[MAILTO] Registering as default handler", { scheme, ok });
  }
  // GNOME's and KDE's "default mail client" covers mailto: links only;
  // opening a saved message is a separate file-type default.
  claimFileTypes({
    // app.setDesktopName (app/index.js) sets CHROME_DESKTOP, the id
    // setAsDefaultProtocolClient uses too.
    desktopId: process.env.CHROME_DESKTOP || "outlook-for-linux.desktop",
    run: runCommand,
    exists: (file) => fs.existsSync(file),
    env: process.env,
  }).then((results) => {
    for (const { mimeType, result } of results) {
      console.info("[MAILTO] Default for file type", { mimeType, result });
    }
  }).catch((err) => console.error("[MAILTO] Could not claim file types", { error: err.message }));
}

// The mailbox the compose window should open on: the one on screen, else the
// remembered one, else wherever the app starts (sign-in for a new profile).
function composeOrigin() {
  const contents = activeContents();
  const current = contents ? mailboxOrigin(contents.getURL()) : null;
  if (current) return current;
  const remembered = accounts && !accounts.getActive().main ? accounts.getActive().mailboxOrigin : store.get(MAILBOX_ORIGIN_KEY);
  return new URL(startUrl(appConfig.url, configOptions.url.default, remembered)).origin;
}

// The visible account's session, which a compose window must share.
function activePartition() {
  return accounts?.getActive().partition ?? appConfig.partition;
}

// Opens Outlook's compose deeplink in a pop-out, as Outlook's own "open in
// new window" does, so the mailbox in the main window keeps its state. With
// the integrated title bar it gets a title bar strip in the main window's
// header colour, as Outlook for Windows gives its pop-outs.
async function openCompose(fields) {
  if (isIntegratedTitleBar(appConfig)) await refreshMainHeaderColor();
  console.info("[MAILTO] Opening compose", {
    recipients: fields.to.length + fields.cc.length + fields.bcc.length,
    hasSubject: Boolean(fields.subject),
    hasBody: Boolean(fields.body),
  });
  const compose = new BrowserWindow({
    title: "Outlook for Linux",
    width: 900,
    height: 760,
    icon: appIcon.windowImage(),
    autoHideMenuBar: true,
    ...titleBarWindowOptions(appConfig, nativeTheme.shouldUseDarkColors),
    webPreferences: mainWebPreferences(appConfig, activePartition()),
  });
  const contents = compose.webContents;
  composeContents.add(contents);
  contents.once("destroyed", () => composeContents.delete(contents));
  applyLinkPolicy(contents, appConfig, activePartition());
  attachPageFeatures(contents);
  contents.loadURL(composeUrl(fields, composeOrigin()), { userAgent: appConfig.chromeUserAgent })
    .catch(() => console.warn("[MAILTO] Compose window failed to load"));
}

// An unhandled rejection ends the app (see app/index.js).
function openComposeFields(fields) {
  openCompose(fields).catch((err) => console.error("[MAILTO] Could not open compose", { error: err.message }));
}

function openComposeSafely(mailto) {
  const fields = parseMailto(mailto);
  if (fields) openComposeFields(fields);
}

function openEmlFile(file) {
  emlViewer.open(file).catch((err) => console.error("[EML] Could not open message", { error: err.message }));
}

/**
 * Handles the command line of this launch or of a second one: a mailto: link
 * opens a compose window, .eml files open in the message viewer, anything
 * else brings the main window back.
 * @param {string[]} args
 * @param {string} [workingDirectory] - where relative paths in args start
 *   (a second launch's own working directory)
 * @returns {boolean} true when a mailto: link or a message was handled
 */
function handleLaunchArgs(args, workingDirectory = process.cwd()) {
  const mailto = findMailtoArg(args);
  const emlFiles = findEmlArgs(args, workingDirectory);
  if (!mailto && !emlFiles.length) return false;
  if (!window) {
    if (mailto) pendingMailto = mailto;
    pendingEmlFiles.push(...emlFiles);
    return true;
  }
  if (mailto) openComposeSafely(mailto);
  for (const file of emlFiles) openEmlFile(file);
  return true;
}

function onAppReady(config, settingsStore) {
  appConfig = config;
  store = settingsStore;
  appIcon = new AppIcon(config, settingsStore, ICON_PATH);
  emlViewer = new EmlViewer({
    openCompose: openComposeFields,
    openMailto: openComposeSafely,
    openExternal: (url) => openOutboundLink(url, config),
    icon: () => appIcon.windowImage(),
  });
  buildMenu();
  registerMailClient(config);
  createWindow(config, settingsStore);
  if (pendingMailto) {
    openComposeSafely(pendingMailto);
    pendingMailto = null;
  }
  for (const file of pendingEmlFiles.splice(0)) openEmlFile(file);
  return window;
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
  handleLaunchArgs,
  activeContents,
  get window() {
    return window;
  },
};
