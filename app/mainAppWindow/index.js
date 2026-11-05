const { app, BrowserWindow, Menu, ipcMain, nativeTheme, shell, session } = require("electron");
const path = require("node:path");
const windowStateKeeper = require("electron-window-state");
const { classifyWindowOpen, classifyNavigation } = require("./linkPolicy");
const { mailboxOrigin, startUrl, isSignedOutPage } = require("./startUrl");
const { PROTOCOL_SCHEMES, parseMailto, findMailtoArg, composeUrl } = require("./mailtoLink");
const configOptions = require("../config/options");
const AppTray = require("../tray");
const UnreadService = require("../unread");
const NotificationService = require("../notifications");
const ConnectionMonitor = require("../connection");
const DownloadFeedback = require("../downloads");
const Zoom = require("../zoom");
const CustomCSS = require("../customCSS");
const AppIcon = require("../appIcon");
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
// A mailto: link that arrived before the main window existed.
let pendingMailto = null;
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
    if (!isMainWindowSender(event) && !isComposeSender(event)) return;
    const target = BrowserWindow.fromWebContents(event.sender);
    if (!target) return;
    const update = sanitizeOverlayUpdate(payload);
    if (!update) return;
    const { sample, ...overlay } = update;
    if (sample) {
      Object.assign(overlay, await samplePageColors(target, update.height));
    }
    if (target === window) refreshMainHeaderColor();
    try {
      if (!target.isDestroyed()) target.setTitleBarOverlay(overlay);
    } catch (err) {
      console.warn("[TITLE_BAR] Could not update the window controls", { error: err.message });
    }
  });
}

// The header colour for compose windows' title bar strips. Not the colour
// behind the main window's controls: that end of the bar holds the avatar.
// The top two rows, across the whole width, are bare header (its icons,
// search box and avatar are vertically centred), or the average of a theme
// image. Sampled when a compose window opens (the main window reports its
// controls' colour mostly while still loading) and on those reports; open
// compose windows follow changes, such as a theme switch.
async function refreshMainHeaderColor() {
  if (!window || window.isDestroyed()) return;
  try {
    const [contentWidth] = window.getContentSize();
    const image = await window.webContents.capturePage({ x: 0, y: 1, width: contentWidth, height: 2 });
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
async function samplePageColors(target, height) {
  try {
    const [contentWidth] = target.getContentSize();
    const image = await target.webContents.capturePage(sampleRect(contentWidth, height));
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
          icon: appIcon.windowImage(),
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
    attachPageFeatures(child.webContents);
  });

  // See classifyNavigation for why main-frame https navigation is allowed.
  webContents.on("will-navigate", (event, url) => {
    const action = classifyNavigation(url);
    if (action === "allow") return;
    event.preventDefault();
    if (action === "external") openExternal(url);
  });
}

// Saved zoom and the user's custom CSS, for every Outlook window.
function attachPageFeatures(webContents) {
  zoom?.attach(webContents);
  customCSS?.attach(webContents);
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
    icon: appIcon.windowImage(),
    autoHideMenuBar: config.menubarAutoHide,
    ...titleBarWindowOptions(config, nativeTheme.shouldUseDarkColors),
    webPreferences: mainWebPreferences(config),
  });
  windowState.manage(window);
  registerIpcHandlers(config);

  if (config.trayIcon) {
    tray = new AppTray(window, appIcon.path, quit);
    buildMenu();
  }
  // The dock badge: Electron sends app.setBadgeCount over Unity
  // LauncherEntry, which GNOME's docks and KDE's task manager show.
  const unread = new UnreadService(() => window, tray);
  unread.attach(session.fromPartition(config.partition));
  const notifications = config.notifications ? new NotificationService(() => window, appIcon.path) : null;
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

  new ConnectionMonitor(window.webContents, config.chromeUserAgent).initialize();
  const downloads = new DownloadFeedback(window, appIcon.path);
  downloads.attach(session.fromPartition(config.partition));
  // Window icons follow in AppIcon itself; the menu's checkbox states too.
  appIcon.onChange((iconPath) => {
    tray?.setBaseIconPath(iconPath);
    notifications?.setIconPath(iconPath);
    downloads.setIconPath(iconPath);
    buildMenu();
  });

  // A new page (reload, sign-out) has no unread count until Outlook reports one.
  window.webContents.on("did-navigate", () => unread.reset());

  configureSpellChecker(config);
  zoom = new Zoom(() => window, settingsStore);
  customCSS = new CustomCSS(config.customCSSLocation);
  attachPageFeatures(window.webContents);
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
    {
      label: "View",
      submenu: [
        { role: "reload" },
        { role: "forceReload" },
        { role: "toggleDevTools" },
        { type: "separator" },
        // Not the zoom roles: these save the zoom for the next launch. The
        // keys are handled by app/zoom (before the page sees them), so the
        // accelerators here are labels only.
        { label: "Actual Size", accelerator: "CmdOrCtrl+0", registerAccelerator: false, click: (_item, win) => zoom?.step(0, win) },
        { label: "Zoom In", accelerator: "CmdOrCtrl+Plus", registerAccelerator: false, click: (_item, win) => zoom?.step(1, win) },
        { label: "Zoom Out", accelerator: "CmdOrCtrl+-", registerAccelerator: false, click: (_item, win) => zoom?.step(-1, win) },
        { type: "separator" },
        { role: "togglefullscreen" },
        { type: "separator" },
        appIcon.menuItem(() => window),
      ],
    },
    { role: "windowMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  // Reachable from the tray too: the integrated title bar has no menu bar.
  tray?.setMenuItems([appIcon.menuItem(() => window)]);
}

function quit() {
  quitting = true;
  app.quit();
}

app.on("before-quit", () => {
  quitting = true;
});

// Opt-in (defaultMailClient): claiming mailto: silently would take it from
// the user's current mail client. Inside Flatpak and Snap, xdg-settings cannot
// change the host's handlers, so the desktop's own settings are the way there;
// the desktop entry already lists the app as a mailto: handler.
function registerMailClient(config) {
  if (!config.defaultMailClient) return;
  if (process.env.FLATPAK_ID || process.env.SNAP_INSTANCE_NAME) {
    console.info("[MAILTO] Sandboxed install: choose the default mail client in the desktop settings");
    return;
  }
  for (const scheme of PROTOCOL_SCHEMES) {
    if (app.isDefaultProtocolClient(scheme)) continue;
    const ok = app.setAsDefaultProtocolClient(scheme);
    console.info("[MAILTO] Registering as default handler", { scheme, ok });
  }
}

// The mailbox the compose window should open on: the one on screen, else the
// remembered one, else wherever the app starts (sign-in for a new profile).
function composeOrigin() {
  const current = window && !window.isDestroyed() ? mailboxOrigin(window.webContents.getURL()) : null;
  if (current) return current;
  return new URL(startUrl(appConfig.url, configOptions.url.default, store.get(MAILBOX_ORIGIN_KEY))).origin;
}

// Opens Outlook's compose deeplink in a pop-out, as Outlook's own "open in
// new window" does, so the mailbox in the main window keeps its state. With
// the integrated title bar it gets a title bar strip in the main window's
// header colour, as Outlook for Windows gives its pop-outs.
async function openCompose(mailto) {
  const fields = parseMailto(mailto);
  if (!fields) return;
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
    webPreferences: mainWebPreferences(appConfig),
  });
  const contents = compose.webContents;
  composeContents.add(contents);
  contents.once("destroyed", () => composeContents.delete(contents));
  applyLinkPolicy(contents, appConfig);
  attachPageFeatures(contents);
  contents.loadURL(composeUrl(fields, composeOrigin()), { userAgent: appConfig.chromeUserAgent })
    .catch(() => console.warn("[MAILTO] Compose window failed to load"));
}

// An unhandled rejection ends the app (see app/index.js).
function openComposeSafely(mailto) {
  openCompose(mailto).catch((err) => console.error("[MAILTO] Could not open compose", { error: err.message }));
}

/**
 * Handles the command line of this launch or of a second one: a mailto: link
 * opens a compose window, anything else brings the main window back.
 * @param {string[]} args
 * @returns {boolean} true when a mailto: link was handled
 */
function handleLaunchArgs(args) {
  const mailto = findMailtoArg(args);
  if (!mailto) return false;
  if (!window) {
    pendingMailto = mailto;
  } else {
    openComposeSafely(mailto);
  }
  return true;
}

function onAppReady(config, settingsStore) {
  appConfig = config;
  store = settingsStore;
  appIcon = new AppIcon(config, settingsStore, ICON_PATH);
  buildMenu();
  registerMailClient(config);
  createWindow(config, settingsStore);
  if (pendingMailto) {
    openComposeSafely(pendingMailto);
    pendingMailto = null;
  }
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
  get window() {
    return window;
  },
};
