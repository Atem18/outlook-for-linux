// Preload for the main Outlook window. Runs sandboxed in an isolated world
// (contextIsolation): it shares Outlook's DOM but not its JavaScript globals.
// Only what tools explicitly pass through contextBridge reaches the page.
//
// A sandboxed preload can only require "electron", so this file and its tools
// are bundled by scripts/buildPreload.js into preload.bundle.js, which is what
// the window loads. Edit this file, not the bundle.
const { contextBridge, ipcRenderer } = require("electron");
const notifications = require("./tools/notifications");
const integratedTitleBar = require("./tools/integratedTitleBar");
const unreadBadge = require("./tools/unreadBadge");

// The subset of the startup config the tools read; see the
// "get-renderer-config" handler in app/mainAppWindow/index.js. Synchronous on
// purpose: the notification shim must be in place before Outlook's scripts
// run, which is before any async reply could arrive.
function readConfig() {
  try {
    return ipcRenderer.sendSync("get-renderer-config") ?? {};
  } catch (err) {
    console.error("[PRELOAD] Could not read the renderer config", { error: err.message });
    return {};
  }
}

const config = readConfig();

function safely(name, fn) {
  try {
    fn();
  } catch (err) {
    console.error(`[PRELOAD] Failed to initialise ${name}`, { error: err.message });
  }
}

// Page-world hooks: before any page script.
safely("notifications", () => notifications.init(config, ipcRenderer, contextBridge));

// DOM tools: once the document exists.
function initDomTools() {
  safely("integratedTitleBar", () => integratedTitleBar.init(config, ipcRenderer));
  safely("unreadBadge", () => unreadBadge.init(config, ipcRenderer));
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initDomTools, { once: true });
} else {
  initDomTools();
}
