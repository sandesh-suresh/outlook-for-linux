const DEFAULT_TITLE = "Microsoft Outlook";
const VALID_URGENCIES = new Set(["low", "normal", "critical"]);

/**
 * Renders native OS notifications on behalf of the Outlook page.
 *
 * The page's own `Notification` calls are intercepted in preload and arrive
 * here over IPC, so notifications survive the page being backgrounded and
 * respect the user's desktop notification settings.
 */
class NotificationService {
  #config;
  #Notification;
  #nativeImage;
  #window = null;

  /**
   * @param {object} config - The resolved startup config.
   * @param {object} [deps] - Electron collaborators, injectable for tests.
   */
  constructor(config, deps = {}) {
    const electron = deps.Notification ? deps : require("electron");
    this.#config = config;
    this.#Notification = electron.Notification;
    this.#nativeImage = electron.nativeImage;
  }

  /** @param {object} window - The main BrowserWindow. */
  setWindow(window) {
    this.#window = window;
  }

  /** @param {object} ipcMain - Electron's ipcMain. */
  initialize(ipcMain) {
    // Renders a native OS notification forwarded from the Outlook page.
    ipcMain.handle("show-notification", (_event, options) =>
      this.show(options),
    );
  }

  /**
   * @param {object} options - `{ id, title, body, icon, urgency }`.
   * @returns {Promise<void>} Always resolves.
   */
  async show(options) {
    if (this.#config.disableNotifications) {
      return;
    }

    const { id, title, body, icon, urgency } = options ?? {};

    try {
      const notification = new this.#Notification({
        title:
          typeof title === "string" && title !== "" ? title : DEFAULT_TITLE,
        body: typeof body === "string" ? body : "",
        icon: this.#loadIcon(icon),
        urgency: VALID_URGENCIES.has(urgency)
          ? urgency
          : this.#config.defaultNotificationUrgency,
      });

      // Both paths report back so the page can clear its own state; guard so a
      // click followed by the platform's close event only reports once.
      let reported = false;
      const report = () => {
        if (reported) {
          return;
        }
        reported = true;
        this.#window?.webContents.send("notification-closed", { id });
      };

      notification.on("click", () => {
        this.#focusWindow();
        report();
      });
      notification.on("close", report);

      notification.show();
      // Never log title or body: they carry sender names and mail subjects.
      console.debug("[NOTIFY] Notification shown");
    } catch (error) {
      console.error("[NOTIFY] Failed to show notification:", error.message);
    }
  }

  /**
   * Only `data:` URLs are supported. Anything else (including remote avatar
   * URLs) is dropped so this path never makes a network request.
   */
  #loadIcon(icon) {
    if (typeof icon !== "string" || !icon.startsWith("data:")) {
      return undefined;
    }
    try {
      const image = this.#nativeImage.createFromDataURL(icon);
      return image.isEmpty() ? undefined : image;
    } catch (error) {
      console.error(
        "[NOTIFY] Failed to decode notification icon:",
        error.message,
      );
      return undefined;
    }
  }

  #focusWindow() {
    try {
      this.#window?.show();
      this.#window?.focus();
    } catch (error) {
      console.error("[NOTIFY] Failed to focus window:", error.message);
    }
  }
}

module.exports = NotificationService;
