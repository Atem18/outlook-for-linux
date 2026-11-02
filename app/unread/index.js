const { app, ipcMain } = require("electron");
const {
  apiOriginOf,
  bearerFrom,
  getFolderUrl,
  getFolderBody,
  parseUnreadCount,
} = require("./outlookApi");
const { sanitizeUnreadUpdate } = require("../tray/badge");

const POLL_INTERVAL_MS = 30_000;
// An API count older than this no longer beats the folder pane's.
const API_FRESH_MS = 2 * POLL_INTERVAL_MS;

/**
 * The Inbox unread count for the tray, from two sources:
 *
 * 1. Outlook's own API: like teams-for-linux's graphApi, which borrows the
 *    Graph token the Teams web app already holds, this borrows the bearer
 *    token the Outlook web app sends with its own /owa/service.svc calls and
 *    asks the same API, read-only, for the Inbox unread count. The token
 *    stays in this process's memory, is never logged or sent to a renderer,
 *    and only ever goes back to the origin it came from.
 * 2. The folder pane (app/browser/tools/unreadBadge.js): the fallback when the
 *    API has not answered recently. Outlook hides the pane in narrow windows.
 *
 * The badge image is drawn by the page, which has a canvas; the main process
 * asks for it with the count to draw.
 */
class UnreadService {
  #getWindow;
  #tray;
  #session = null;
  #token = null; // { origin, authorization }
  #apiCount = null;
  #apiAt = 0;
  #domCount = null;
  #shown = -1;
  #timer = null;
  #inFlight = false;

  /**
   * @param {() => Electron.BrowserWindow | null} getWindow
   * @param {import("../tray")} tray
   */
  constructor(getWindow, tray) {
    this.#getWindow = getWindow;
    this.#tray = tray;
  }

  /** @param {Electron.Session} session - the Outlook session */
  attach(session) {
    this.#session = session;
    // Electron keeps one onBeforeSendHeaders listener per session; this is
    // the only one. A later module needing it (e.g. Intune SSO) must share it.
    session.webRequest.onBeforeSendHeaders(
      { urls: ["https://*/owa/service.svc*"] },
      (details, callback) => {
        this.#capture(details);
        callback({ requestHeaders: details.requestHeaders });
      }
    );

    // Inbox count read from the folder pane, when the pane is rendered
    ipcMain.on("unread-dom-count", (event, count) => {
      if (!this.#isFromMainWindow(event) || !Number.isInteger(count) || count < 0) return;
      this.#domCount = count;
      this.#update();
    });
    // Badge image the page drew for the count it was asked to draw
    ipcMain.on("unread-badge", (event, payload) => {
      if (!this.#isFromMainWindow(event)) return;
      const update = sanitizeUnreadUpdate(payload);
      if (update && update.count === this.#shown) this.#tray?.setUnread(update.count, update.badge);
    });

    this.#timer = setInterval(() => this.#poll(), POLL_INTERVAL_MS);
  }

  /** A new page (reload, sign-out, account switch) starts from nothing. */
  reset() {
    this.#token = null;
    this.#apiCount = null;
    this.#apiAt = 0;
    this.#domCount = null;
    this.#show(0);
  }

  destroy() {
    clearInterval(this.#timer);
  }

  #isFromMainWindow(event) {
    const window = this.#getWindow();
    return window && !window.isDestroyed() && event.sender === window.webContents;
  }

  #capture(details) {
    const origin = apiOriginOf(details.url);
    if (!origin) return;
    const authorization = bearerFrom(details.requestHeaders);
    if (!authorization) return;
    const first = !this.#token;
    this.#token = { origin, authorization };
    if (first) setTimeout(() => this.#poll(), 0);
  }

  async #poll() {
    const token = this.#token;
    if (!token || !this.#session || this.#inFlight) return;
    this.#inFlight = true;
    try {
      const response = await this.#session.fetch(getFolderUrl(token.origin), {
        method: "POST",
        headers: {
          authorization: token.authorization,
          action: "GetFolder",
          "content-type": "application/json; charset=utf-8",
        },
        body: getFolderBody(),
      });
      if (response.status === 401 || response.status === 403) {
        // Expired: forget it; Outlook's next call brings a fresh one.
        if (this.#token === token) this.#token = null;
        return;
      }
      if (!response.ok) {
        console.warn("[UNREAD] Outlook API request failed", { status: response.status });
        return;
      }
      const count = parseUnreadCount(await response.json());
      if (count === null) {
        console.warn("[UNREAD] Outlook API response had no unread count");
        return;
      }
      this.#apiCount = count;
      this.#apiAt = Date.now();
      this.#update();
    } catch (err) {
      console.warn("[UNREAD] Outlook API request error", { error: err.message });
    } finally {
      this.#inFlight = false;
    }
  }

  #update() {
    const apiFresh = this.#apiCount !== null && Date.now() - this.#apiAt < API_FRESH_MS;
    const count = apiFresh ? this.#apiCount : this.#domCount;
    if (count !== null) this.#show(count);
  }

  #show(count) {
    if (count === this.#shown) return;
    this.#shown = count;
    app.setBadgeCount(count);
    if (count === 0) {
      this.#tray?.setUnread(0, null);
      return;
    }
    const window = this.#getWindow();
    if (window && !window.isDestroyed()) {
      window.webContents.send("unread-render", count);
    }
  }
}

module.exports = UnreadService;
