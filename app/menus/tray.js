/**
 * Owns the system tray icon and its context menu.
 *
 * Electron collaborators are injected so this module can be unit tested
 * outside a running Electron process.
 */
class ApplicationTray {
  #window;
  #config;
  #iconPath;
  #Tray;
  #Menu;
  #nativeImage;
  #tray = null;
  #unreadCount = 0;

  /**
   * @param {object} window - The main BrowserWindow.
   * @param {object} config - The resolved startup config.
   * @param {string} iconPath - Absolute path to the base tray icon.
   * @param {object} [deps] - Electron collaborators, injectable for tests.
   */
  constructor(window, config, iconPath, deps = {}) {
    const electron = deps.Tray ? deps : require("electron");
    this.#window = window;
    this.#config = config;
    this.#iconPath = iconPath;
    this.#Tray = electron.Tray;
    this.#Menu = electron.Menu;
    this.#nativeImage = electron.nativeImage;
  }

  /** @param {object} ipcMain - Electron's ipcMain. */
  initialize(ipcMain) {
    if (!this.#config.trayIconEnabled) {
      console.info("[TRAY] Tray icon disabled by config");
      return;
    }

    try {
      this.#tray = new this.#Tray(
        this.#nativeImage.createFromPath(this.#iconPath),
      );
      this.#tray.setToolTip(this.#config.appTitle);
      this.#tray.setContextMenu(
        this.#Menu.buildFromTemplate(this.#buildMenu()),
      );
      this.#tray.on("click", () => this.#toggleWindow());
      console.info("[TRAY] Tray created");
    } catch (error) {
      console.error("[TRAY] Failed to create tray:", error.message);
      return;
    }

    ipcMain.on("tray-update", (_event, payload) =>
      this.updateTrayImage(payload),
    );
  }

  #buildMenu() {
    return [
      { label: this.#unreadLabel(), click: () => this.#toggleWindow() },
      { type: "separator" },
      { label: "Show / Hide", click: () => this.#toggleWindow() },
      { label: "Reload", click: () => this.#window.reload?.() },
      { type: "separator" },
      { label: "Quit", click: () => this.#window.emit?.("tray-quit") },
    ];
  }

  #unreadLabel() {
    if (this.#unreadCount <= 0) {
      return "No unread emails";
    }
    if (this.#unreadCount === 1) {
      return "1 unread email";
    }
    return `${this.#unreadCount} unread emails`;
  }

  #refreshMenu() {
    if (!this.#tray) {
      return;
    }
    this.#tray.setContextMenu(this.#Menu.buildFromTemplate(this.#buildMenu()));
  }

  #toggleWindow() {
    if (this.#window.isVisible()) {
      this.#window.hide();
      return;
    }
    this.#window.show();
    this.#window.focus();
  }

  /**
   * @param {{icon: string|null, flash: boolean, count: number}} payload
   */
  updateTrayImage(payload) {
    if (!this.#tray) {
      return;
    }

    try {
      const { icon = null, flash = false, count = 0 } = payload ?? {};

      // A null icon means "no badge": fall back to the icon on disk rather
      // than leaving a stale badged image in the tray.
      this.#tray.setImage(
        icon
          ? this.#nativeImage.createFromDataURL(icon)
          : this.#nativeImage.createFromPath(this.#iconPath),
      );
      this.#tray.setToolTip(
        count > 0
          ? `${this.#config.appTitle} (${count})`
          : this.#config.appTitle,
      );
      this.#window.flashFrame(Boolean(flash));

      this.#unreadCount = count;
      this.#refreshMenu();
    } catch (error) {
      console.error("[TRAY] Failed to update tray image:", error.message);
    }
  }

  close() {
    if (!this.#tray) {
      return;
    }
    this.#tray.destroy();
    this.#tray = null;
  }
}

module.exports = ApplicationTray;
