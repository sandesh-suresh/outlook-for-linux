const { app, ipcMain } = require("electron");
const fs = require("node:fs");
const path = require("node:path");

const { isNetworkError } = require("./utils/networkErrors");
const { AppConfiguration } = require("./appConfiguration");
const CommandLineManager = require("./startup/commandLine");
const TrayIconChooser = require("./browser/tools/trayIconChooser");
const ApplicationTray = require("./menus/tray");
const NotificationService = require("./notifications/service");
const { createWindow } = require("./mainAppWindow");
const { createGuardedIpcMain, registerIpcHandlers } = require("./ipc/register");

// Registered before anything else so a failure during startup is logged
// rather than lost. A transient network error must not take the app down:
// Electron surfaces those as uncaught exceptions from the renderer's own
// load failures, and killing the process would look like a random crash to
// the user.
process.on("uncaughtException", (error) => {
  if (isNetworkError(error.message)) {
    console.warn("[NETWORK] Recoverable network error:", error.message);
    return;
  }
  console.error("[FATAL] Uncaught exception:", error.message, error.stack);
  process.exit(1);
});

process.on("unhandledRejection", (reason) => {
  const message = reason?.message ?? String(reason);
  if (isNetworkError(message)) {
    console.warn("[NETWORK] Recoverable network rejection:", message);
    return;
  }
  console.error("[FATAL] Unhandled rejection:", message);
  process.exit(1);
});

// A second instance would fight the first over the tray icon and the single
// Outlook session, so hand off to the running one instead.
if (!app.requestSingleInstanceLock()) {
  console.info("[STARTUP] Another instance is already running, exiting");
  app.quit();
  return;
}

const appConfiguration = new AppConfiguration(
  app.getPath("userData"),
  app.getVersion(),
);
const config = appConfiguration.startupConfig;

CommandLineManager.applySwitches(config, app);

let window = null;
let tray = null;

app.on("second-instance", () => {
  window?.show();
  window?.focus();
});

app.on("window-all-closed", () => {
  tray?.close();
  app.quit();
});

app.whenReady().then(() => {
  const guardedIpcMain = createGuardedIpcMain(ipcMain);
  const notificationService = new NotificationService(config);

  registerIpcHandlers(guardedIpcMain, {
    config,
    setBadgeCount: (count) => app.setBadgeCount(count),
  });
  notificationService.initialize(guardedIpcMain);

  window = createWindow(config);
  notificationService.setWindow(window);

  tray = new ApplicationTray(
    window,
    config,
    new TrayIconChooser(config).getFile(),
  );
  tray.initialize(guardedIpcMain);

  // The tray's Quit item destroys rather than closes: closeToTray installs a
  // close handler that hides the window, which would otherwise make Quit a
  // no-op. destroy() skips that handler.
  window.on("tray-quit", () => {
    tray?.close();
    window?.destroy();
    app.quit();
  });

  if (config.watchConfigFile && config.isConfigFile) {
    watchConfigFile(appConfiguration.configPath);
  }

  console.info("[STARTUP] Application ready");
});

/**
 * Warns when the config file changes. Config is immutable after startup, so
 * this only tells the user a restart is needed — it never re-reads.
 */
function watchConfigFile(configPath) {
  try {
    fs.watch(
      path.join(configPath, "config.json"),
      { persistent: false },
      () => {
        console.info("[CONFIG] Config file changed; restart to apply");
      },
    );
  } catch (error) {
    console.error("[CONFIG] Failed to watch config file:", error.message);
  }
}
