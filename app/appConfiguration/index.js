const { default: Store } = require("electron-store");

/**
 * Owns the startup config (immutable after load) and the persistent settings
 * store. Other modules receive what they need from here rather than loading
 * config themselves.
 */
class AppConfiguration {
  #configPath;
  #startupConfig;
  #settingsStore;

  /**
   * @param {string} configPath - Directory holding config.json (userData)
   * @param {string} appVersion - Current application version
   */
  constructor(configPath, appVersion) {
    this.#configPath = configPath;
    this.#startupConfig = Object.freeze(require("../config")(configPath, appVersion));
    this.#settingsStore = new Store({ name: "settings", clearInvalidConfig: true });
  }

  get configPath() {
    return this.#configPath;
  }

  get configFilePath() {
    return require("../config").getConfigFilePath(this.#configPath);
  }

  get startupConfig() {
    return this.#startupConfig;
  }

  get settingsStore() {
    return this.#settingsStore;
  }
}

module.exports = { AppConfiguration };
