const { Notification, nativeImage, shell } = require("electron");
const path = require("node:path");

/**
 * Feedback for attachment downloads. Electron already asks where to save; this
 * adds the window's progress bar and a notification when the file is saved,
 * which opens it on click. File names and paths are never logged.
 */
class DownloadFeedback {
  #window;
  #icon;
  #active = new Set();
  // Electron's Notification is only weakly held: without a reference, a
  // click from the notification history does nothing (teams-for-linux #3121).
  #notifications = new Set();

  /**
   * @param {Electron.BrowserWindow} window
   * @param {string} iconPath
   */
  constructor(window, iconPath) {
    this.#window = window;
    this.setIconPath(iconPath);
  }

  /** @param {string} iconPath - the icon for notifications from now on */
  setIconPath(iconPath) {
    this.#icon = nativeImage.createFromPath(iconPath).resize({ width: 128, height: 128 });
  }

  /** @param {Electron.Session} session - the Outlook session */
  attach(session) {
    session.on("will-download", (_event, item) => this.#track(item));
  }

  #track(item) {
    this.#active.add(item);
    item.on("updated", () => this.#updateProgress());
    item.once("done", (_event, state) => {
      this.#active.delete(item);
      this.#updateProgress();
      if (state === "completed") {
        this.#notifyDone(item.getSavePath());
      } else if (state === "interrupted") {
        console.warn("[DOWNLOAD] Download interrupted");
      }
    });
  }

  // One bar for all running downloads; indeterminate when sizes are unknown.
  #updateProgress() {
    if (this.#window.isDestroyed()) return;
    if (this.#active.size === 0) {
      this.#window.setProgressBar(-1);
      return;
    }
    let received = 0;
    let total = 0;
    for (const item of this.#active) {
      received += item.getReceivedBytes();
      total += item.getTotalBytes();
    }
    this.#window.setProgressBar(total > 0 ? received / total : 2);
  }

  #notifyDone(savePath) {
    if (!savePath || !Notification.isSupported()) return;
    const notification = new Notification({
      title: "Download complete",
      body: path.basename(savePath),
      icon: this.#icon,
    });
    this.#notifications.add(notification);
    const release = () => this.#notifications.delete(notification);
    notification.on("click", () => {
      release();
      shell.openPath(savePath).then((error) => {
        if (error) console.warn("[DOWNLOAD] Could not open the downloaded file");
      });
    });
    notification.on("close", release);
    notification.show();
  }
}

module.exports = DownloadFeedback;
