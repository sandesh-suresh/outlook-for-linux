// Configuration option definitions for Outlook for Linux.
//
// Single source of truth for the wrapper's config surface. Consumed at runtime
// by the yargs parser in ./index.js. Keep it a plain data module with no
// imports so it stays loadable outside Electron.

module.exports = {
  url: {
    default: "https://outlook.office.com",
    describe:
      "Outlook web URL. Use https://outlook.live.com for personal accounts, or a tenant-specific URL.",
    type: "string",
  },
  appTitle: {
    default: "Microsoft Outlook",
    describe: "Window and tray tooltip title",
    type: "string",
  },
  appIcon: {
    default: "",
    describe:
      "Absolute path to a custom tray icon. Empty means use the bundled icon.",
    type: "string",
  },
  appIconType: {
    default: "default",
    describe: "Which bundled tray icon variant to use",
    type: "string",
    choices: ["default", "light", "dark"],
  },
  trayIconEnabled: {
    default: true,
    describe: "Enable the tray icon",
    type: "boolean",
  },
  closeToTray: {
    default: true,
    describe: "Hide to tray instead of quitting when the window is closed",
    type: "boolean",
  },
  disableNotifications: {
    default: false,
    describe: "Disable all native notifications",
    type: "boolean",
  },
  defaultNotificationUrgency: {
    default: "normal",
    describe: "Urgency passed to the native notification server",
    type: "string",
    choices: ["low", "normal", "critical"],
  },
  useUnreadListLogic: {
    default: true,
    describe:
      "Watch the mail list with a MutationObserver to derive the unread count for the tray badge",
    type: "boolean",
  },
  disableBadgeCount: {
    default: false,
    describe: "Do not draw the unread count badge on the tray icon",
    type: "boolean",
  },
  disableNotificationWindowFlash: {
    default: false,
    describe: "Do not flash the window frame when the unread count increases",
    type: "boolean",
  },
  disableGpu: {
    default: false,
    describe: "Disable GPU hardware acceleration",
    type: "boolean",
  },
  logConfig: {
    default: {
      transports: {
        console: { level: "info" },
        file: { level: false },
      },
    },
    describe:
      'electron-log configuration. Set to the string "console" for plain console logging, or false to disable logging.',
    type: "object",
  },
  watchConfigFile: {
    default: false,
    describe:
      "Log a warning when the config file changes. Config is read once at startup, so a restart is required to apply changes.",
    type: "boolean",
  },
  webDebug: {
    default: false,
    describe: "Open DevTools on start",
    type: "boolean",
  },
};
