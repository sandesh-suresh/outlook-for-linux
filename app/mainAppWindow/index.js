const path = require("node:path");

const DEFAULT_WIDTH = 1280;
const DEFAULT_HEIGHT = 800;

// Hosts that are part of the Outlook experience or its sign-in flow. Sending
// any of these to the system browser would break authentication.
const IN_APP_HOST_SUFFIXES = [
  "outlook.office.com",
  "outlook.office365.com",
  "outlook.live.com",
  "office.com",
  "login.microsoftonline.com",
  "login.live.com",
  "login.microsoft.com",
  "microsoftonline.com",
];

/**
 * Decides whether a navigation target belongs in the system browser.
 *
 * @param {string} url - The target URL.
 * @param {string} appUrl - The configured Outlook URL.
 * @returns {boolean} True when the OS should handle it.
 */
function shouldOpenExternally(url, appUrl) {
  let target;
  try {
    target = new URL(url);
  } catch {
    // Unparseable input is never handed to the shell.
    return false;
  }

  if (target.protocol !== "http:" && target.protocol !== "https:") {
    // mailto:, tel: and friends are the OS's business.
    return true;
  }

  let appHost;
  try {
    appHost = new URL(appUrl).hostname;
  } catch {
    appHost = "";
  }

  const host = target.hostname;
  const isInApp =
    host === appHost ||
    IN_APP_HOST_SUFFIXES.some(
      (suffix) => host === suffix || host.endsWith(`.${suffix}`),
    );

  return !isInApp;
}

/**
 * @param {object} config - The resolved startup config.
 * @param {object} [deps] - Electron collaborators, injectable for tests.
 * @returns {object} The created window.
 */
function createWindow(config, deps = {}) {
  const electron = deps.BrowserWindow ? deps : require("electron");
  const { BrowserWindow, shell } = electron;
  const preload =
    deps.preloadPath ?? path.join(__dirname, "..", "browser", "preload.js");

  const window = new BrowserWindow({
    width: DEFAULT_WIDTH,
    height: DEFAULT_HEIGHT,
    title: config.appTitle,
    autoHideMenuBar: true,
    webPreferences: {
      preload,
      // The preload script needs `require` to reach Electron's nativeImage for
      // tray badge rendering, so it is not sandboxed and context isolation is
      // off. The IPC allowlist in app/security/ipcValidator.js is the
      // compensating control: the page can only reach the channels listed
      // there, with payloads sanitised.
      sandbox: false,
      contextIsolation: false,
      nodeIntegration: false,
      spellcheck: false,
    },
  });

  window.on("close", (event) => {
    // Only hide when there is a tray to restore from; hiding without one
    // would leave the app running with no way to get it back.
    if (config.closeToTray && config.trayIconEnabled) {
      event.preventDefault();
      window.hide();
    }
  });

  window.webContents.setWindowOpenHandler(({ url }) => {
    if (shouldOpenExternally(url, config.url)) {
      shell.openExternal(url);
      return { action: "deny" };
    }
    return { action: "allow" };
  });

  if (config.webDebug) {
    window.webContents.openDevTools();
  }

  window.loadURL(config.url);
  console.info("[WINDOW] Main window created");

  return window;
}

module.exports = { createWindow, shouldOpenExternally };
