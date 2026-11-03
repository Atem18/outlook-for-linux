const { nextZoomFactor, validZoomFactor } = require("./steps");

const STORE_KEY = "zoomFactor";

/**
 * Remembers the zoom across restarts (Ctrl+Plus / Ctrl+Minus / Ctrl+0, the
 * View menu and Ctrl+wheel), the way teams-for-linux app/browser/tools/zoom.js
 * does, but from the main process so no IPC is needed. Chromium applies a
 * zoom to every window of the same origin, so pop-outs follow the main window.
 */
class Zoom {
  #getWindow;
  #settingsStore;

  /**
   * @param {() => Electron.BrowserWindow|null} getWindow
   * @param {{get: Function, set: Function}} settingsStore
   */
  constructor(getWindow, settingsStore) {
    this.#getWindow = getWindow;
    this.#settingsStore = settingsStore;
  }

  get factor() {
    return validZoomFactor(this.#settingsStore.get(STORE_KEY));
  }

  /**
   * Applies the saved zoom to each page the window loads (sign-in and Outlook
   * are different origins, and a reload resets it) and follows Ctrl+wheel.
   * @param {Electron.WebContents} webContents
   */
  attach(webContents) {
    webContents.on("did-finish-load", () => webContents.setZoomFactor(this.factor));
    webContents.on("zoom-changed", (_event, direction) => {
      const { BrowserWindow } = require("electron");
      this.step(direction === "in" ? 1 : -1, BrowserWindow.fromWebContents(webContents));
    });
  }

  /**
   * @param {1|-1|0} direction
   * @param {Electron.BrowserWindow|null} [window] - the window the menu or
   *   wheel acted on; the main window by default
   */
  step(direction, window = this.#getWindow()) {
    if (!window || window.isDestroyed()) return;
    const factor = nextZoomFactor(window.webContents.getZoomFactor(), direction);
    window.webContents.setZoomFactor(factor);
    this.#settingsStore.set(STORE_KEY, factor);
  }
}

module.exports = Zoom;
