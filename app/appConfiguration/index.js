const argv = require("../config");

/**
 * Owns the application configuration for the process lifetime.
 *
 * The resolved config is read once at startup and kept behind a private field.
 * It is already deep-frozen by the config loader; treat it as immutable.
 */
class AppConfiguration {
  #configPath;
  #startupConfig;

  /**
   * @param {string} configPath - Directory holding config.json.
   * @param {string} appVersion - Version reported by --version.
   * @param {string[]} [argvArray] - Argument vector; defaults to the real one.
   */
  constructor(configPath, appVersion, argvArray) {
    this.#configPath = configPath;
    this.#startupConfig = argv(configPath, appVersion, argvArray);
  }

  get configPath() {
    return this.#configPath;
  }

  get startupConfig() {
    return this.#startupConfig;
  }
}

module.exports = { AppConfiguration };
