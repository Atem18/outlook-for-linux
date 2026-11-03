const { Tray, Menu, nativeImage } = require("electron");
const { trayTooltip, isBadgeDataUrl, compositeBadge } = require("./badge");

const ICON_SIZE = 64;

/**
 * System tray icon: shows the unread count, toggles the window on click, and
 * offers Show and Quit. Adapted from teams-for-linux app/menus/tray.js.
 */
class AppTray {
  #tray;
  #window;
  #baseImage;
  #onQuit;
  #lastCount = 0;
  #lastBadge = null;

  /**
   * @param {Electron.BrowserWindow} window
   * @param {string} iconPath
   * @param {() => void} onQuit - quits the app for real (bypasses close-to-tray)
   */
  constructor(window, iconPath, onQuit) {
    this.#window = window;
    this.#onQuit = onQuit;
    this.#baseImage = AppTray.#loadBase(iconPath);
    this.#tray = new Tray(this.#baseImage);
    this.#tray.setToolTip(trayTooltip(0));
    this.#tray.on("click", () => this.toggleWindow());
    this.setMenuItems([]);
  }

  /**
   * Rebuilds the context menu with extra items between Show and Quit (the
   * App Icon submenu: with the integrated title bar there is no menu bar).
   * @param {Electron.MenuItemConstructorOptions[]} items
   */
  setMenuItems(items) {
    if (this.#tray.isDestroyed()) return;
    this.#tray.setContextMenu(
      Menu.buildFromTemplate([
        { label: "Show Outlook", click: () => this.showWindow() },
        { type: "separator" },
        ...(items.length ? [...items, { type: "separator" }] : []),
        { label: "Quit", click: () => this.#onQuit() },
      ])
    );
  }

  // Tray icons are drawn at 22-24px; a 64px source scales down cleanly.
  static #loadBase(iconPath) {
    return nativeImage.createFromPath(iconPath).resize({ width: ICON_SIZE, height: ICON_SIZE });
  }

  /**
   * Swaps the icon (App Icon menu), keeping the unread badge on it.
   * @param {string} iconPath
   */
  setBaseIconPath(iconPath) {
    if (this.#tray.isDestroyed()) return;
    this.#baseImage = AppTray.#loadBase(iconPath);
    this.setUnread(this.#lastCount, this.#lastBadge);
  }

  showWindow() {
    if (this.#window.isDestroyed()) return;
    if (this.#window.isMinimized()) this.#window.restore();
    this.#window.show();
    this.#window.focus();
  }

  toggleWindow() {
    if (this.#window.isDestroyed()) return;
    if (this.#window.isVisible() && this.#window.isFocused()) {
      this.#window.hide();
    } else {
      this.showWindow();
    }
  }

  /**
   * @param {number} count - unread count
   * @param {string|null} badgeDataUrl - the count drawn on a transparent
   *   64x64 PNG by the page (the main process has no canvas); null for none
   */
  setUnread(count, badgeDataUrl) {
    if (this.#tray.isDestroyed()) return;
    const image = count > 0 ? this.#badgedImage(badgeDataUrl) : null;
    this.#tray.setImage(image ?? this.#baseImage);
    this.#tray.setToolTip(trayTooltip(count));
    // Ask for attention only when the count goes up, not on every redraw or
    // when mail is read.
    if (count > this.#lastCount && !this.#window.isFocused()) {
      this.#window.flashFrame(true);
    }
    this.#lastCount = count;
    this.#lastBadge = badgeDataUrl;
  }

  #badgedImage(badgeDataUrl) {
    if (!isBadgeDataUrl(badgeDataUrl)) return null;
    // A malformed payload yields an empty image rather than throwing.
    let badge = nativeImage.createFromDataURL(badgeDataUrl);
    if (badge.isEmpty()) return null;
    const { width, height } = badge.getSize();
    if (width !== ICON_SIZE || height !== ICON_SIZE) {
      badge = badge.resize({ width: ICON_SIZE, height: ICON_SIZE });
    }
    const pixels = compositeBadge(this.#baseImage.toBitmap(), badge.toBitmap());
    if (!pixels) return null;
    return nativeImage.createFromBitmap(pixels, { width: ICON_SIZE, height: ICON_SIZE });
  }

  destroy() {
    if (!this.#tray.isDestroyed()) this.#tray.destroy();
  }
}

module.exports = AppTray;
