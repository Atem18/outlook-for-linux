const { app } = require("electron");
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
  // A second launch (desktop icon, taskbar) focuses the running window.
  // mailto: arguments will be routed from here once the handler exists.
  app.on("second-instance", () => mainAppWindow.show());

  app.on("window-all-closed", () => app.quit());

  app.whenReady().then(() => mainAppWindow.onAppReady(config, appConfig.settingsStore));
} else {
  console.info("[STARTUP] Another instance is running; handing over");
  app.quit();
}
