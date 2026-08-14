const {
  validateIpcChannel,
  allowedChannels,
} = require("../security/ipcValidator");

const MAX_BADGE_COUNT = 9999;

/**
 * Wraps ipcMain so no channel can be registered or served without passing the
 * allowlist.
 *
 * Registering an unlisted channel throws immediately: a missing allowlist
 * entry is a programming error, and failing at startup is far easier to
 * diagnose than a handler that silently never fires.
 *
 * @param {object} ipcMain - Electron's ipcMain.
 * @returns {{on: Function, handle: Function}} The guarded façade.
 */
function createGuardedIpcMain(ipcMain) {
  const assertAllowed = (channel) => {
    if (!allowedChannels.has(channel)) {
      throw new Error(
        `IPC channel "${channel}" is not allowlisted. Add it to app/security/ipcValidator.js.`,
      );
    }
  };

  const guard = (channel, handler) => (event, payload) => {
    if (!validateIpcChannel(channel, payload)) {
      return undefined;
    }
    return handler(event, payload);
  };

  return {
    on(channel, handler) {
      assertAllowed(channel);
      ipcMain.on(channel, guard(channel, handler));
    },
    handle(channel, handler) {
      assertAllowed(channel);
      const guarded = guard(channel, handler);
      // `invoke()` in the renderer always yields a promise, so the registered
      // handler presents one too, whether or not it is itself async. A
      // synchronous handler that returned a bare value would still work in
      // Electron, but callers here can rely on the promise contract.
      ipcMain.handle(channel, async (event, payload) =>
        guarded(event, payload),
      );
    },
  };
}

/**
 * Registers the channels that belong to the main process itself. Tray and
 * notification channels are registered by their own modules, through the same
 * guarded façade.
 *
 * @param {object} guardedIpcMain - From createGuardedIpcMain.
 * @param {object} deps - `{ config, setBadgeCount }`.
 */
function registerIpcHandlers(guardedIpcMain, { config, setBadgeCount }) {
  // Returns the resolved startup config to the renderer.
  guardedIpcMain.handle("get-config", () => config);

  // Sets the OS taskbar/dock badge to the unread count.
  guardedIpcMain.handle("set-badge-count", (_event, count) => {
    if (!Number.isInteger(count) || count < 0 || count > MAX_BADGE_COUNT) {
      console.warn("[IPC] Ignoring out-of-range badge count");
      return;
    }
    try {
      setBadgeCount(count);
    } catch (error) {
      // Not every desktop environment supports badges; a failure here must
      // not surface as a rejected invoke in the page.
      console.debug("[IPC] Badge count not applied:", error.message);
    }
  });

  // Renderer-side uncaught error, forwarded from preload.
  guardedIpcMain.on("window-error", (_event, payload) => {
    console.error(
      `[RENDERER] ${payload?.message ?? "unknown error"} (${payload?.source ?? "?"}:${payload?.line ?? "?"})`,
    );
  });

  // Renderer-side unhandled promise rejection, forwarded from preload.
  guardedIpcMain.on("unhandled-rejection", (_event, payload) => {
    console.error(
      `[RENDERER] Unhandled rejection: ${payload?.reason ?? "unknown"}`,
    );
  });
}

module.exports = { createGuardedIpcMain, registerIpcHandlers };
