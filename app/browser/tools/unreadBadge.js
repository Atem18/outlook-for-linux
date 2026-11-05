/**
 * Unread badge, renderer half.
 *
 * - Reports the Inbox unread count shown in Outlook's folder pane to the main
 *   process (app/unread), which uses it when Outlook's API has not answered.
 * - Draws the badge for the count the main process settles on (the main
 *   process has no canvas) and sends it back for the tray.
 *
 * Why the folder pane: Outlook does call the Badging API
 * (navigator.setAppBadge), but only for mail that arrives while it is in the
 * background, and it calls clearAppBadge() on every load and whenever the
 * window is focused, so it is not the unread total. Electron sends those
 * calls straight to the dock (the same LauncherEntry badge as
 * app.setBadgeCount), where they overwrote the real count, so the Badging
 * API is taken away from the page (blockPageBadge) and Outlook, which checks
 * for it, leaves the dock alone. The folder pane's Inbox item carries the total, under a
 * data-folder-name that does not depend on the interface language.
 *
 * Narrow windows: Outlook does not render the folder pane at all below a
 * certain width. The last count read is then kept rather than reset.
 */

const DOM_COUNT_CHANNEL = "unread-dom-count";
const RENDER_CHANNEL = "unread-render";
const BADGE_CHANNEL = "unread-badge";
const SIZE = 64;
const CHECK_DELAY_MS = 1000;
const INBOX_SELECTOR = '[role="treeitem"][data-folder-name="inbox" i]';

/**
 * The Inbox unread count, 0 when Outlook shows none, or null when the folder
 * pane is not rendered (narrow window, loading, sign-in page).
 * @param {ParentNode} root
 * @returns {number|null}
 */
function readInboxUnread(root = document) {
  const inbox = root.querySelector(INBOX_SELECTOR);
  if (!inbox) return null;
  for (const el of inbox.querySelectorAll("span")) {
    if (el.children.length > 0) continue;
    const text = el.textContent.trim();
    if (/^\d+$/.test(text)) return Number(text);
  }
  return 0;
}

// Text drawn on the badge: counts above 99 would not fit a tray icon.
function badgeLabel(count) {
  return count > 99 ? "99+" : String(count);
}

// A red disc with the count in the bottom-right corner, on a transparent
// 64x64 canvas; the main process composites it over the app icon.
function renderBadge(count) {
  const canvas = document.createElement("canvas");
  canvas.width = SIZE;
  canvas.height = SIZE;
  const ctx = canvas.getContext("2d");
  const label = badgeLabel(count);
  const radius = label.length > 2 ? 26 : 22;
  const cx = SIZE - radius;
  const cy = SIZE - radius;
  ctx.fillStyle = "#d13438";
  ctx.beginPath();
  ctx.arc(cx, cy, radius, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = "#ffffff";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.font = `bold ${label.length > 2 ? 22 : 30}px sans-serif`;
  ctx.fillText(label, cx, cy + 2);
  return canvas.toDataURL("image/png");
}

// Runs in the page's world, before Outlook's scripts. Neither the
// disable-blink-features switch nor webPreferences.disableBlinkFeatures
// removes the Badging API in Electron 44, so it is deleted here.
function removeBadgingApi() {
  for (const name of ["setAppBadge", "clearAppBadge"]) {
    delete Navigator.prototype[name];
  }
}

class UnreadBadge {
  #ipcRenderer = null;
  #last = -1;
  #scheduled = false;

  /**
   * The folder-pane count feeds the dock badge too, so it runs without the
   * tray; badge images are only drawn for the tray.
   * @param {object} config - renderer config (`trayIcon` flag)
   * @param {Electron.IpcRenderer} ipcRenderer
   */
  init(config, ipcRenderer) {
    if (!ipcRenderer) return;
    this.#ipcRenderer = ipcRenderer;
    if (config?.trayIcon) ipcRenderer.on(RENDER_CHANNEL, (_event, count) => this.#render(count));
    // Counts change through text updates inside the pane, so character data
    // is watched too; checks are batched to one per second.
    new MutationObserver(() => this.#schedule()).observe(document.documentElement, {
      childList: true,
      subtree: true,
      characterData: true,
    });
    this.#check();
  }

  /** @param {Electron.ContextBridge} contextBridge */
  blockPageBadge(contextBridge) {
    contextBridge.executeInMainWorld({ func: removeBadgingApi });
  }

  #schedule() {
    if (this.#scheduled) return;
    this.#scheduled = true;
    setTimeout(() => {
      this.#scheduled = false;
      this.#check();
    }, CHECK_DELAY_MS);
  }

  #check() {
    const count = readInboxUnread();
    if (count === null || count === this.#last) return;
    this.#last = count;
    this.#ipcRenderer.send(DOM_COUNT_CHANNEL, count);
  }

  #render(count) {
    if (!Number.isInteger(count) || count <= 0) return;
    let badge = null;
    try {
      badge = renderBadge(count);
    } catch (err) {
      console.debug("[UNREAD] Badge render failed", { error: err.message });
    }
    this.#ipcRenderer.send(BADGE_CHANNEL, { count, badge });
  }
}

module.exports = new UnreadBadge();
module.exports._internals = { readInboxUnread, badgeLabel, removeBadgingApi, INBOX_SELECTOR };
