const { app, ipcMain } = require("electron");
const RequestHeaders = require("../requestHeaders");
const {
  apiOriginOf,
  bearerFrom,
  getFolderUrl,
  getFolderBody,
  parseUnreadCount,
} = require("./outlookApi");
const { sanitizeUnreadUpdate, unreadTotal } = require("../tray/badge");

const POLL_INTERVAL_MS = 30_000;
// An API count older than this no longer beats the folder pane's.
const API_FRESH_MS = 2 * POLL_INTERVAL_MS;

/**
 * The Inbox unread count of one account, from two sources:
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
 * One per account (each has its own session and page); UnreadTotal below
 * adds them up for the dock badge and the tray.
 */
class UnreadService {
  #getContents;
  #onCount;
  #session = null;
  #token = null; // { origin, authorization }
  #apiCount = null;
  #apiAt = 0;
  #domCount = null;
  #shown = -1;
  #timer = null;
  #inFlight = false;
  #domListener = null;
  #destroyed = false;

  /**
   * @param {() => Electron.WebContents | null} getContents - the account's
   *   Outlook page
   * @param {(count: number) => void} onCount - called when the count changes
   */
  constructor(getContents, onCount) {
    this.#getContents = getContents;
    this.#onCount = onCount;
  }

  /** @param {Electron.Session} session - the account's session */
  attach(session) {
    this.#session = session;
    // Electron keeps one onBeforeSendHeaders listener per session, so it is
    // shared through app/requestHeaders (Intune SSO uses it too).
    RequestHeaders.for(session).use({
      name: "unread",
      urls: ["https://*/owa/service.svc*"],
      matches: (url) => !this.#destroyed && apiOriginOf(url) !== null,
      handle: (details) => this.#capture(details),
    });

    // Inbox count read from the folder pane, when the pane is rendered
    this.#domListener = (event, count) => {
      if (!this.#isFromOwnPage(event) || !Number.isInteger(count) || count < 0) return;
      this.#domCount = count;
      this.#update();
    };
    ipcMain.on("unread-dom-count", this.#domListener);

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
    this.#destroyed = true;
    this.#token = null;
    clearInterval(this.#timer);
    if (this.#domListener) ipcMain.removeListener("unread-dom-count", this.#domListener);
  }

  #isFromOwnPage(event) {
    const contents = this.#getContents();
    return Boolean(contents) && !contents.isDestroyed() && event.sender === contents;
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
    this.#onCount(count);
  }
}

/**
 * Adds up the accounts' unread counts for the dock badge (Electron sends
 * app.setBadgeCount over Unity LauncherEntry, which GNOME's docks and KDE's
 * task manager show) and the tray. The tray's badge image is drawn by a page,
 * which has a canvas; the main process asks for it with the count to draw.
 */
class UnreadTotal {
  #tray;
  #getRenderer;
  #onCounts;
  #counts = new Map();
  #shown = -1;

  /**
   * @param {import("../tray")|null} tray - null without a tray icon
   * @param {() => Electron.WebContents | null} getRenderer - the page that
   *   draws the tray badge (the main window's)
   * @param {(counts: Record<string, number>) => void} [onCounts] - every
   *   account's count, on any change (the account switcher shows them)
   */
  constructor(tray, getRenderer, onCounts = () => {}) {
    this.#tray = tray;
    this.#getRenderer = getRenderer;
    this.#onCounts = onCounts;
    // Badge image the page drew for the count it was asked to draw
    ipcMain.on("unread-badge", (event, payload) => {
      const renderer = this.#getRenderer();
      if (!renderer || event.sender !== renderer) return;
      const update = sanitizeUnreadUpdate(payload);
      if (update && update.count === this.#shown) this.#tray?.setUnread(update.count, update.badge);
    });
  }

  /**
   * @param {string} key - the account id
   * @param {number} count
   */
  set(key, count) {
    this.#counts.set(key, count);
    this.#update();
  }

  /** @param {string} key - an account that is gone */
  delete(key) {
    if (this.#counts.delete(key)) this.#update();
  }

  get counts() {
    return Object.fromEntries(this.#counts);
  }

  #update() {
    this.#onCounts(this.counts);
    const total = unreadTotal(this.#counts.values());
    if (total === this.#shown) return;
    this.#shown = total;
    app.setBadgeCount(total);
    if (total === 0) {
      this.#tray?.setUnread(0, null);
      return;
    }
    const renderer = this.#getRenderer();
    if (this.#tray && renderer && !renderer.isDestroyed()) renderer.send("unread-render", total);
  }
}

module.exports = { UnreadService, UnreadTotal };
