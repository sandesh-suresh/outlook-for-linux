const { ipcRenderer } = require("electron");
const { installNotificationShim } = require("./tools/notificationShim");
const mutationTitle = require("./tools/mutationTitle");
const trayIconRenderer = require("./tools/trayIconRenderer");

// Installed synchronously, before any page script runs. Waiting for the
// config round-trip would let Outlook capture the real Notification
// constructor first, and notifications would bypass the shim entirely.
installNotificationShim(globalThis, ipcRenderer);

// Surface renderer-side failures in the main log, where users find them.
globalThis.addEventListener("error", (event) => {
  ipcRenderer.send("window-error", {
    message: event.message,
    source: event.filename,
    line: event.lineno,
  });
});

globalThis.addEventListener("unhandledrejection", (event) => {
  ipcRenderer.send("unhandled-rejection", {
    reason: String(event.reason?.message ?? event.reason),
  });
});

globalThis.addEventListener("DOMContentLoaded", async () => {
  try {
    const config = await ipcRenderer.invoke("get-config");

    // Both are config-gated and DOM-dependent, so they wait for the config
    // round-trip; the notification shim above deliberately does not.
    mutationTitle.init(config);
    if (config.trayIconEnabled) {
      trayIconRenderer.init(config, ipcRenderer);
    }
    console.debug("[PRELOAD] Browser tools initialised");
  } catch (error) {
    console.error("[PRELOAD] Initialisation failed:", error.message);
  }
});
