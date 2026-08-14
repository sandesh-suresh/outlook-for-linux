const path = require("node:path");

const iconFolder = path.join(__dirname, "../..", "assets/icons");

const icons = {
  default: "icon-96x96.png",
  light: "icon-monochrome-light-96x96.png",
  dark: "icon-monochrome-dark-96x96.png",
};

/** Resolves which tray icon file to use, from config. */
class TrayIconChooser {
  #config;

  /** @param {object} config - The resolved startup config. */
  constructor(config) {
    this.#config = config;
  }

  /** @returns {string} An absolute path to the icon file. */
  getFile() {
    const custom = this.#config.appIcon;
    if (typeof custom === "string" && custom.trim() !== "") {
      return custom;
    }
    const file = icons[this.#config.appIconType] ?? icons.default;
    return path.join(iconFolder, file);
  }
}

module.exports = TrayIconChooser;
