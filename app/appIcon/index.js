const { app, BrowserWindow, dialog, nativeImage } = require("electron");
const fs = require("node:fs/promises");
const path = require("node:path");
const { writeConfigOption } = require("../config");
const { customIconPath, windowIconResize } = require("./iconPath");
const themeIcon = require("./themeIcon");

const OPTION = "appIcon";
// The chosen icon is copied here, so it survives the original being moved or
// deleted, and a Flatpak keeps it although the file dialog's grant to the
// original is temporary.
const CHOSEN_ICON_FILE = "app-icon.png";

/**
 * The app icon: the bundled one, or a custom PNG from the appIcon option.
 * The App Icon menu sets that option in config.json, as teams-for-linux's
 * does in its config store, so there is one place for it. Covers the
 * windows, the tray and notifications; outside the Snap it can also copy the
 * icon into the user icon theme so the taskbar and launcher follow. Ported
 * from teams-for-linux app/menus/index.js (chooseAppIcon, resetAppIcon,
 * toggleThemeIcon).
 *
 * Paths are never logged: they hold the user name.
 */
class AppIcon {
  #chosen;
  #defaultPath;
  #path;
  #listeners = [];

  /**
   * @param {{appIcon?: string}} config - the startup config
   * @param {{get: Function, delete: Function}} store - settings store, only
   *   to move an icon the menu kept there before it wrote config.json
   * @param {string} defaultPath - the bundled icon
   */
  constructor(config, store, defaultPath) {
    this.#chosen = AppIcon.#migrate(config.appIcon, store);
    this.#defaultPath = defaultPath;
    this.#path = this.#resolve();
  }

  // Earlier builds kept the menu's choice in settings.json, where it won
  // over the option. Move it into config.json once.
  static #migrate(configured, store) {
    const legacy = store.get(OPTION);
    if (typeof legacy !== "string") return configured;
    try {
      writeConfigOption(app.getPath("userData"), OPTION, legacy);
      store.delete(OPTION);
      console.info("[APP_ICON] Moved the chosen icon from settings.json to config.json");
    } catch (err) {
      console.warn("[APP_ICON] Could not move the chosen icon to config.json", { error: err.name });
    }
    return legacy;
  }

  /** The icon file in use. */
  get path() {
    return this.#path;
  }

  get isCustom() {
    return this.#path !== this.#defaultPath;
  }

  /** The icon sized for a window. */
  windowImage() {
    const image = nativeImage.createFromPath(this.#path);
    const resize = windowIconResize(image.getSize());
    return resize ? image.resize(resize) : image;
  }

  /** @param {(path: string) => void} listener - called when the icon changes */
  onChange(listener) {
    this.#listeners.push(listener);
  }

  /**
   * The App Icon submenu.
   * @param {() => Electron.BrowserWindow|null} getWindow - parent for dialogs
   */
  menuItem(getWindow) {
    return {
      label: "App Icon",
      submenu: [
        { label: "Choose App Icon…", click: () => this.#run(() => this.choose(getWindow())) },
        {
          label: "Also use it for the taskbar and launcher",
          type: "checkbox",
          visible: themeIcon.supported(),
          enabled: this.isCustom,
          checked: this.isCustom && themeIcon.installed(),
          click: () => this.#run(() => this.toggleThemeIcon(getWindow())),
        },
        { label: "Reset to default", enabled: this.isCustom, click: () => this.#run(() => this.reset(getWindow())) },
      ],
    };
  }

  async choose(parent) {
    const { canceled, filePaths } = await dialog.showOpenDialog(parent, {
      title: "Choose App Icon",
      filters: [{ name: "Images", extensions: ["png"] }],
      properties: ["openFile"],
    });
    if (canceled || !filePaths.length) return;
    const [file] = filePaths;
    const image = nativeImage.createFromPath(file);
    if (image.isEmpty()) {
      await dialog.showMessageBox(parent, {
        type: "error",
        title: "Choose App Icon",
        message: "That file could not be read as an image.",
        detail: file,
      });
      return;
    }
    const copy = path.join(app.getPath("userData"), CHOSEN_ICON_FILE);
    await fs.writeFile(copy, image.toPNG());
    await this.#save(parent, copy);
    this.#chosen = copy;
    this.#changed();
    if (themeIcon.supported()) await this.#askThemeIcon(parent, image);
    this.#emit();
  }

  async #askThemeIcon(parent, image) {
    const { response } = await dialog.showMessageBox(parent, {
      type: "question",
      title: "App Icon",
      message: "Also use it for the taskbar and launcher?",
      detail:
        "This copies the icon into your icon theme under ~/.local/share/icons. " +
        "It shows on the next start, on some desktops after you log in again. " +
        "You can change this later under App Icon.",
      buttons: ["Yes", "No"],
      defaultId: 0,
      cancelId: 1,
    });
    if (response === 0) {
      themeIcon.install(image);
    } else {
      themeIcon.remove();
    }
  }

  async toggleThemeIcon(parent) {
    if (themeIcon.installed()) {
      themeIcon.remove();
    } else if (this.isCustom && themeIcon.install(nativeImage.createFromPath(this.#path))) {
      await dialog.showMessageBox(parent, {
        type: "info",
        title: "App Icon",
        message: "The taskbar and launcher pick up the icon on the next start.",
        detail: "Some desktops only refresh it after you log in again.",
      });
    }
    this.#emit();
  }

  async reset(parent) {
    await this.#save(parent, undefined);
    this.#chosen = null;
    this.#changed();
    themeIcon.remove();
    this.#emit();
  }

  // The icon applies for this session either way; only keeping it for the
  // next start needs config.json, which is left alone when it is not valid
  // JSON.
  async #save(parent, value) {
    try {
      writeConfigOption(app.getPath("userData"), OPTION, value);
    } catch (err) {
      console.warn("[APP_ICON] Could not write config.json", { error: err.name });
      await dialog.showMessageBox(parent, {
        type: "warning",
        title: "App Icon",
        message: "The icon applies until the app quits: config.json could not be updated.",
        detail: `Fix the file (${err.message}), then choose the icon again.`,
      });
    }
  }

  #resolve() {
    const custom = customIconPath(this.#chosen);
    if (!custom) return this.#defaultPath;
    if (nativeImage.createFromPath(custom).isEmpty()) {
      console.warn("[APP_ICON] The custom icon could not be read, using the bundled one");
      return this.#defaultPath;
    }
    return custom;
  }

  #changed() {
    this.#path = this.#resolve();
    for (const win of BrowserWindow.getAllWindows()) {
      if (!win.isDestroyed()) win.setIcon(this.windowImage());
    }
  }

  #emit() {
    for (const listener of this.#listeners) listener(this.#path);
  }

  // Menu clicks cannot await; an unhandled rejection ends the app.
  #run(action) {
    action().catch((err) => console.error("[APP_ICON] Changing the icon failed", { error: err.message }));
  }
}

module.exports = AppIcon;
