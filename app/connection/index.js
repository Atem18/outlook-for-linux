const { net, powerMonitor } = require("electron");
const { isNetworkError, retryDelay } = require("./retryPolicy");

/**
 * Reloads Outlook when it failed to load for lack of network: with backoff
 * while offline, and at once on resume from suspend. Outlook handles short
 * outages itself once loaded; this covers the page never loading at all.
 */
class ConnectionMonitor {
  #webContents;
  #userAgent;
  #failedUrl = null;
  #attempt = 0;
  #timer = null;
  #errorPagePending = false;

  /**
   * @param {Electron.WebContents} webContents - the main window's
   * @param {string} userAgent
   */
  constructor(webContents, userAgent) {
    this.#webContents = webContents;
    this.#userAgent = userAgent;
  }

  initialize() {
    this.#webContents.on("did-fail-load", (_event, errorCode, _description, url, isMainFrame) => {
      if (!isMainFrame || !isNetworkError(errorCode)) return;
      // The failed URL is kept to reload it, never logged: it can carry tokens.
      this.#failedUrl = url;
      this.#errorPagePending = true;
      console.warn("[CONNECTION] Outlook failed to load for lack of network", { errorCode, attempt: this.#attempt });
      this.#schedule(retryDelay(this.#attempt));
    });
    this.#webContents.on("did-finish-load", () => {
      // A failed load is followed by did-finish-load for Chromium's error
      // page, while getURL() still reports Outlook's URL; only a later, real
      // page load clears the failure.
      if (this.#errorPagePending) {
        this.#errorPagePending = false;
        return;
      }
      this.#reset();
    });
    powerMonitor.on("resume", () => {
      if (this.#failedUrl) this.#schedule(0);
    });
    this.#webContents.once("destroyed", () => clearTimeout(this.#timer));
  }

  #schedule(delay) {
    clearTimeout(this.#timer);
    this.#timer = setTimeout(() => this.#retry(), delay);
  }

  #retry() {
    if (!this.#failedUrl || this.#webContents.isDestroyed()) return;
    this.#attempt += 1;
    if (!net.isOnline()) {
      this.#schedule(retryDelay(this.#attempt));
      return;
    }
    console.info("[CONNECTION] Network is back, reloading Outlook", { attempt: this.#attempt });
    this.#webContents.loadURL(this.#failedUrl, { userAgent: this.#userAgent }).catch(() => {});
  }

  #reset() {
    if (this.#failedUrl) console.info("[CONNECTION] Outlook loaded again");
    clearTimeout(this.#timer);
    this.#failedUrl = null;
    this.#attempt = 0;
  }
}

module.exports = ConnectionMonitor;
