const { app, session } = require("electron");
const path = require("node:path");
const log = require("electron-log/main");

// Pinned so development (`electron .`, named after package.json "name") and
// packaged builds (named after productName) share ~/.config/outlook-for-linux.
// An explicit --user-data-dir wins: it is how a second account runs as its own
// instance, with its own config, session and single-instance lock.
if (!app.commandLine.hasSwitch("user-data-dir")) {
  app.setPath("userData", path.join(app.getPath("appData"), "outlook-for-linux"));
}

// Route console.* through electron-log so main-process logs carry levels and
// timestamps. File logging stays at its default location in userData/logs.
log.initialize();
Object.assign(console, log.functions);

// Show "Outlook for Linux" in notifications and the desktop environment.
if (process.platform === "linux") {
  app.setName("Outlook for Linux");

  // Sandboxed installs export the desktop entry under their own id.
  if (process.env.FLATPAK_ID) {
    app.setDesktopName(`${process.env.FLATPAK_ID}.desktop`);
  } else if (process.env.SNAP_INSTANCE_NAME) {
    app.setDesktopName(`${process.env.SNAP_INSTANCE_NAME}_outlook-for-linux.desktop`);
  } else {
    app.setDesktopName("outlook-for-linux.desktop");
  }
}

process.on("uncaughtException", (error) => {
  console.error("[FATAL] Uncaught exception:", { message: error?.message, stack: error?.stack });
  process.exit(1);
});

process.on("unhandledRejection", (reason) => {
  const message = reason instanceof Error ? reason.message : String(reason);
  console.error("[FATAL] Unhandled promise rejection:", { message });
  process.exit(1);
});

// E2E tests run against a throwaway userData directory for a clean state.
if (process.env.E2E_USER_DATA_DIR) {
  app.setPath("userData", process.env.E2E_USER_DATA_DIR);
}

const { AppConfiguration } = require("./appConfiguration");
const appConfig = new AppConfiguration(app.getPath("userData"), app.getVersion());
const config = appConfig.startupConfig;

const mainAppWindow = require("./mainAppWindow");

if (app.requestSingleInstanceLock()) {
  // A second launch with a mailto: link (another app's "send email") opens a
  // compose window; any other launch (desktop icon, taskbar) focuses the
  // running window.
  app.on("second-instance", (_event, args) => {
    if (!mainAppWindow.handleLaunchArgs(args)) mainAppWindow.show();
  });

  app.on("window-all-closed", () => app.quit());

  // A mailto: link on this launch's own command line waits for the window.
  mainAppWindow.handleLaunchArgs(process.argv);
  app.whenReady().then(async () => {
    // Intune first: the first sign-in request already needs the broker's
    // cookie. Waits up to ~5 s for the broker, only when enabled.
    if (config.intune) {
      const intune = require("./intune");
      await intune.initSso(config.intuneUser);
      intune.attach(session.fromPartition(config.partition));
    }
    const clientCertificate = require("./clientCertificate");
    if (config.clientCertPinDialog) clientCertificate.initialize();
    clientCertificate.importPkcs12(config.clientCertPath, config.clientCertPassword);
    mainAppWindow.onAppReady(config, appConfig.settingsStore);
    require("./globalShortcuts").register(config, () => mainAppWindow.window, app);
    if (process.platform === "linux" && config.webauthn) {
      require("./webauthn")
        .initialize(mainAppWindow.window, config)
        .catch((err) => console.error("[WEBAUTHN] Could not initialise", { error: err.message }));
    }
  });
} else {
  console.info("[STARTUP] Another instance is running; handing over");
  app.quit();
}
