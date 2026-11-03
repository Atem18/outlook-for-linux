const fs = require("node:fs/promises");

/**
 * Injects the user's own stylesheet (customCSSLocation) into Outlook pages.
 * Ported from teams-for-linux app/customCSS. Teams' packaged themes restyle
 * Teams and do not apply here; Outlook's own themes and dark mode (which
 * follows the desktop's light/dark preference) cover theming.
 */
class CustomCSS {
  #location;
  #css = null;

  /** @param {string} location - path to a .css file, or "" for none */
  constructor(location) {
    this.#location = typeof location === "string" ? location.trim() : "";
  }

  get enabled() {
    return this.#location !== "";
  }

  // Read once: the file is the user's, and a missing or unreadable one only
  // means no custom style. Log the failure, not the path (it can hold the
  // user name).
  async #read() {
    if (this.#css === null) {
      try {
        this.#css = await fs.readFile(this.#location, "utf8");
      } catch (err) {
        console.warn("[CUSTOM_CSS] Could not read the custom CSS file", { code: err.code });
        this.#css = "";
      }
    }
    return this.#css;
  }

  /** @param {Electron.WebContents} webContents */
  attach(webContents) {
    if (!this.enabled) return;
    webContents.on("did-finish-load", async () => {
      const css = await this.#read();
      if (css && !webContents.isDestroyed()) webContents.insertCSS(css).catch(() => {});
    });
  }
}

module.exports = CustomCSS;
