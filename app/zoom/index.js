const { nextZoomFactor, validZoomFactor, zoomDirectionForInput } = require("./steps");

const STORE_KEY = "zoomFactor";

// Electron reports one Ctrl+wheel notch as two zoom-changed events about
// 1.5 ms apart (seen on Electron 44, Linux), which skipped a preset per
// notch. A real wheel never notches this fast.
const DUPLICATE_WHEEL_MS = 10;

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
   * are different origins, and a reload resets it), and follows the zoom keys
   * and Ctrl+wheel.
   *
   * The keys are read in before-input-event rather than left to the menu
   * accelerators: those only see keys the page did not consume, they miss
   * the numpad and Ctrl+0 on layouts where 0 needs Shift, and pop-outs have
   * no menu of their own. preventDefault here also stops the accelerator, so
   * a key zooms once.
   * @param {Electron.WebContents} webContents
   */
  attach(webContents) {
    const { BrowserWindow } = require("electron");
    webContents.on("did-finish-load", () => webContents.setZoomFactor(this.factor));
    webContents.on("before-input-event", (event, input) => {
      const direction = zoomDirectionForInput(input);
      if (direction === null) return;
      event.preventDefault();
      this.step(direction, BrowserWindow.fromWebContents(webContents));
    });
    let lastWheel = { direction: null, at: 0 };
    webContents.on("zoom-changed", (_event, zoomDirection) => {
      const now = performance.now();
      const duplicate = zoomDirection === lastWheel.direction && now - lastWheel.at < DUPLICATE_WHEEL_MS;
      lastWheel = { direction: zoomDirection, at: now };
      if (duplicate) return;
      this.step(zoomDirection === "in" ? 1 : -1, BrowserWindow.fromWebContents(webContents));
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
