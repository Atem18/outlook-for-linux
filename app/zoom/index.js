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
 * zoom per origin and session; the new factor is applied to every attached
 * page, so pop-outs and the other accounts' pages follow at once.
 */
class Zoom {
  #getContents;
  #settingsStore;
  #attached = new Set();

  /**
   * @param {() => Electron.WebContents|null} getContents - the page the menu
   *   acts on by default (the visible account's)
   * @param {{get: Function, set: Function}} settingsStore
   */
  constructor(getContents, settingsStore) {
    this.#getContents = getContents;
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
    this.#attached.add(webContents);
    webContents.once("destroyed", () => this.#attached.delete(webContents));
    webContents.on("did-finish-load", () => webContents.setZoomFactor(this.factor));
    webContents.on("before-input-event", (event, input) => {
      const direction = zoomDirectionForInput(input);
      if (direction === null) return;
      event.preventDefault();
      this.step(direction, webContents);
    });
    let lastWheel = { direction: null, at: 0 };
    webContents.on("zoom-changed", (_event, zoomDirection) => {
      const now = performance.now();
      const duplicate = zoomDirection === lastWheel.direction && now - lastWheel.at < DUPLICATE_WHEEL_MS;
      lastWheel = { direction: zoomDirection, at: now };
      if (duplicate) return;
      this.step(zoomDirection === "in" ? 1 : -1, webContents);
    });
  }

  /**
   * @param {1|-1|0} direction
   * @param {Electron.WebContents|null} [contents] - the page the keys or
   *   wheel acted on; the visible account's by default
   */
  step(direction, contents = this.#getContents()) {
    if (!contents || contents.isDestroyed()) return;
    const factor = nextZoomFactor(contents.getZoomFactor(), direction);
    for (const page of this.#attached) {
      if (!page.isDestroyed()) page.setZoomFactor(factor);
    }
    contents.setZoomFactor(factor);
    this.#settingsStore.set(STORE_KEY, factor);
  }
}

module.exports = Zoom;
