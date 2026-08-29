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
    // about:blank is how a same-origin popup bootstraps itself before
    // Outlook's own script navigates it to real content — unlike mailto:,
    // tel: and true external protocol handlers, it is never the OS's
    // business, or the popup ends up as an empty tab in the system browser.
    if (target.protocol === "about:") {
      return false;
    }
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
 * Builds the native right-click menu for a given Chromium context-menu
 * event, since Electron shows no context menu at all unless the app builds
 * one itself from these params — unlike a regular browser, where Chromium
 * supplies one automatically (and already varies it by editable vs.
 * read-only context, which is the "two kinds of menu" a browser shows).
 *
 * @param {object} params - The context-menu event's params (isEditable,
 *   selectionText, editFlags, ...).
 * @returns {object[]} An electron Menu template; empty when there's nothing
 *   actionable to show.
 */
function buildContextMenuTemplate(params) {
  const flags = params.editFlags || {};

  if (params.isEditable) {
    return [
      { label: "Undo", role: "undo", enabled: !!flags.canUndo },
      { label: "Redo", role: "redo", enabled: !!flags.canRedo },
      { type: "separator" },
      { label: "Cut", role: "cut", enabled: !!flags.canCut },
      { label: "Copy", role: "copy", enabled: !!flags.canCopy },
      { label: "Paste", role: "paste", enabled: !!flags.canPaste },
      { type: "separator" },
      { label: "Select All", role: "selectAll", enabled: !!flags.canSelectAll },
    ];
  }

  if (params.selectionText) {
    return [{ label: "Copy", role: "copy", enabled: !!flags.canCopy }];
  }

  return [];
}

/**
 * A popup window created via setWindowOpenHandler's "allow" path (e.g.
 * Outlook's own "Open in a new window") gets a brand-new, fully independent
 * webContents — it does not inherit the opener's context-menu handler, so
 * this must be attached to it separately.
 *
 * @param {object} webContents - The webContents to attach to.
 * @param {object} Menu - The electron Menu class.
 * @param {object} window - The BrowserWindow that owns webContents.
 */
function attachContextMenu(webContents, Menu, window) {
  webContents.on("context-menu", (_event, params) => {
    const template = buildContextMenuTemplate(params);
    if (template.length > 0) {
      Menu.buildFromTemplate(template).popup({ window });
    }
  });
}

/**
 * @param {object} config - The resolved startup config.
 * @param {object} [deps] - Electron collaborators, injectable for tests.
 * @returns {object} The created window.
 */
function createWindow(config, deps = {}) {
  const electron = deps.BrowserWindow ? deps : require("electron");
  const { BrowserWindow, shell, Menu } = electron;
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

  // Outlook's web client sniffs the UA to decide whether it's running in a
  // browser it recognises; when it isn't, it falls back to a generic edit
  // menu instead of its own reply/forward/delete-aware context menu.
  // Electron's default UA always fails that check, since it injects the
  // app name and an "Electron/x.y.z" token that no real browser sends.
  // Stripping just those two tokens, and leaving the OS and Chrome-version
  // tokens exactly as Electron already set them, is enough for Outlook to
  // treat this as a supported Chrome and enable its own menu.
  window.webContents.setUserAgent(
    window.webContents
      .getUserAgent()
      .replace(/outlook-for-linux\/\S+\s+/, "")
      .replace(/\s+Electron\/\S+/, ""),
  );

  window.on("close", (event) => {
    // Only hide when there is a tray to restore from; hiding without one
    // would leave the app running with no way to get it back.
    if (config.closeToTray && config.trayIconEnabled) {
      event.preventDefault();
      window.hide();
    }
  });

  attachContextMenu(window.webContents, Menu, window);

  window.webContents.on("did-create-window", (childWindow) => {
    attachContextMenu(childWindow.webContents, Menu, childWindow);
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

module.exports = { createWindow, shouldOpenExternally, buildContextMenuTemplate };
