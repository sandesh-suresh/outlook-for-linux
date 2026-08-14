/**
 * Replaces the page's `Notification` constructor so Outlook's new-mail
 * notifications are rendered by the main process as native OS notifications
 * instead of by Chromium.
 *
 * Chromium's own notifications inside a wrapped web app do not respect the
 * desktop's notification settings and vanish when the page is backgrounded,
 * which is the behaviour users report as "notifications don't work".
 */
function installNotificationShim(target, ipcRenderer) {
  const live = new Map();
  let nextId = 1;

  class ShimNotification {
    #id;
    #listeners = { click: [], close: [] };
    #closed = false;

    constructor(title, options = {}) {
      this.#id = nextId++;
      this.title = typeof title === "string" ? title : "";
      this.body = typeof options.body === "string" ? options.body : "";
      this.icon = options.icon;
      this.onclick = null;
      this.onclose = null;

      live.set(this.#id, this);

      // Fire-and-forget: the page's Notification constructor is synchronous,
      // so a rejection here has nowhere to go but a log line.
      Promise.resolve(
        ipcRenderer.invoke("show-notification", {
          id: this.#id,
          title: this.title,
          body: this.body,
          icon: this.icon,
          urgency: options.urgency,
        }),
      ).catch((error) => {
        // Never log title or body: they carry sender names and mail subjects.
        console.error(
          "[NOTIFY] Failed to forward notification:",
          error.message,
        );
        live.delete(this.#id);
      });
    }

    get id() {
      return this.#id;
    }

    addEventListener(type, handler) {
      if (this.#listeners[type] && typeof handler === "function") {
        this.#listeners[type].push(handler);
      }
    }

    removeEventListener(type, handler) {
      const handlers = this.#listeners[type];
      if (!handlers) {
        return;
      }
      const index = handlers.indexOf(handler);
      if (index !== -1) {
        handlers.splice(index, 1);
      }
    }

    close() {
      this.#dispatch("close");
    }

    /** Called from the main-process reply. */
    handleRemoteClose() {
      this.#dispatch("close");
    }

    #dispatch(type) {
      if (this.#closed) {
        return;
      }
      this.#closed = true;
      live.delete(this.#id);

      try {
        this[`on${type}`]?.call(this, { type, target: this });
        for (const handler of this.#listeners[type]) {
          handler.call(this, { type, target: this });
        }
      } catch (error) {
        console.error(`[NOTIFY] ${type} handler threw:`, error.message);
      }
    }

    static requestPermission(callback) {
      // The user already granted permission by installing a mail client.
      if (typeof callback === "function") {
        callback("granted");
      }
      return Promise.resolve("granted");
    }

    static get permission() {
      return "granted";
    }
  }

  ipcRenderer.on("notification-closed", (_event, payload) => {
    live.get(payload?.id)?.handleRemoteClose();
  });

  target.Notification = ShimNotification;
}

module.exports = { installNotificationShim };
