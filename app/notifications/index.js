const { Notification, ipcMain, nativeImage } = require("electron");
const { sanitizeNotification, isNotificationId } = require("./sanitize");

// Electron does not always emit close, so live notifications are capped,
// dropping the oldest first (as in teams-for-linux notifications/service.js).
const MAX_TRACKED_NOTIFICATIONS = 50;

/**
 * Shows Outlook's web notifications as native desktop notifications.
 *
 * The page's `Notification` is replaced by a shim (app/browser/tools/
 * notifications.js) that forwards each one here. Clicking a notification
 * brings the window back and relays the click to the page, so Outlook opens
 * the message or reminder it was about. With several accounts, each
 * account's page sends its own; a click shows that account, and the title
 * names it.
 */
class NotificationService {
  #getWindow;
  #isOutlookPage;
  #accountOf;
  #icon;
  // Electron's Notification is only weakly held by V8. Once the JS object is
  // collected, clicking it from the notification history does nothing
  // (teams-for-linux issue #3121), so live ones are kept here by id.
  #live = new Map();

  /**
   * @param {() => Electron.BrowserWindow | null} getWindow
   * @param {string} iconPath
   * @param {object} [accounts] - multiple accounts mode
   * @param {(contents: Electron.WebContents) => boolean} [accounts.isOutlookPage] -
   *   whether a sender is an account's Outlook page; the main window's only
   *   by default
   * @param {(contents: Electron.WebContents) => {name: string|null, show: () => void}|null} [accounts.accountOf] -
   *   the sender's account: the name to add to the title (null with a single
   *   account) and how to make it the visible one
   */
  constructor(getWindow, iconPath, accounts = {}) {
    this.#getWindow = getWindow;
    this.#isOutlookPage = accounts.isOutlookPage ?? ((contents) => contents === this.#getWindow()?.webContents);
    this.#accountOf = accounts.accountOf ?? (() => null);
    this.setIconPath(iconPath);
  }

  /** @param {string} iconPath - the icon for notifications from now on */
  setIconPath(iconPath) {
    this.#icon = nativeImage.createFromPath(iconPath).resize({ width: 128, height: 128 });
  }

  initialize() {
    // Electron loads libnotify at run time; without it every notification is
    // dropped, so say so once instead of failing silently.
    if (!Notification.isSupported()) {
      console.warn("[NOTIFICATIONS] Native notifications are not supported here (is libnotify installed?)");
    }
    // Show a native notification for one the Outlook page created
    ipcMain.on("notification-show", (event, payload) => this.#show(event, payload));
    // Close a notification the Outlook page closed
    ipcMain.on("notification-close", (event, id) => this.#close(event, id));
  }

  #isFromOutlook(event) {
    const window = this.#getWindow();
    return Boolean(window) && !window.isDestroyed() && this.#isOutlookPage(event.sender);
  }

  #show(event, payload) {
    if (!this.#isFromOutlook(event) || !Notification.isSupported()) return;
    const options = sanitizeNotification(payload);
    if (!options) return;

    const account = this.#accountOf(event.sender);
    const notification = new Notification({
      title: account?.name ? `${options.title} (${account.name})` : options.title,
      body: options.body,
      icon: this.#icon,
      silent: options.silent,
      timeoutType: options.requireInteraction ? "never" : "default",
    });

    // The payload sent back is the opaque id and the event type only.
    const sender = event.sender;
    const relay = (type) => {
      if (!sender.isDestroyed()) sender.send("notification-event", { id: options.id, type });
    };

    notification.on("click", () => {
      this.#live.delete(options.id);
      this.#showWindow();
      account?.show();
      relay("click");
    });
    notification.on("close", () => {
      this.#live.delete(options.id);
      relay("close");
    });
    notification.on("failed", () => {
      this.#live.delete(options.id);
      relay("error");
    });

    this.#track(options.id, notification);
    notification.show();
    relay("show");
  }

  #close(event, id) {
    if (!this.#isFromOutlook(event) || !isNotificationId(id)) return;
    const notification = this.#live.get(id);
    if (!notification) return;
    this.#live.delete(id);
    notification.close();
  }

  #track(id, notification) {
    this.#live.set(id, notification);
    while (this.#live.size > MAX_TRACKED_NOTIFICATIONS) {
      this.#live.delete(this.#live.keys().next().value);
    }
  }

  #showWindow() {
    const window = this.#getWindow();
    if (!window || window.isDestroyed()) return;
    if (window.isMinimized()) window.restore();
    window.show();
    window.focus();
  }
}

module.exports = NotificationService;
