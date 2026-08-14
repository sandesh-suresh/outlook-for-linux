const test = require("node:test");
const assert = require("node:assert");
const { createWindow, shouldOpenExternally } = require("../../app/mainAppWindow");

function fakeDeps() {
  const opened = [];
  const created = [];
  class FakeBrowserWindow {
    constructor(options) {
      this.options = options;
      this.loadedUrl = null;
      this.listeners = {};
      this.hidden = false;
      this.webContents = {
        handlers: {},
        setWindowOpenHandler(handler) {
          this.handlers.windowOpen = handler;
        },
        openDevTools() {
          this.devToolsOpen = true;
        },
      };
      created.push(this);
    }
    loadURL(url) {
      this.loadedUrl = url;
    }
    on(event, handler) {
      this.listeners[event] = handler;
    }
    hide() {
      this.hidden = true;
    }
  }
  return {
    created,
    opened,
    deps: {
      BrowserWindow: FakeBrowserWindow,
      shell: { openExternal: (url) => opened.push(url) },
      preloadPath: "/app/browser/preload.js",
    },
  };
}

const CONFIG = {
  url: "https://outlook.office.com",
  appTitle: "Microsoft Outlook",
  closeToTray: true,
  trayIconEnabled: true,
  webDebug: false,
};

test("loads the configured Outlook URL with the preload script attached", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);

  assert.strictEqual(created[0].loadedUrl, "https://outlook.office.com");
  assert.strictEqual(
    created[0].options.webPreferences.preload,
    "/app/browser/preload.js",
  );
  assert.strictEqual(created[0].options.title, "Microsoft Outlook");
});

test("respects a custom URL from config", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, url: "https://outlook.live.com/mail/0/" }, deps);
  assert.strictEqual(created[0].loadedUrl, "https://outlook.live.com/mail/0/");
});

test("opens dev tools only when web debugging is enabled", () => {
  const off = fakeDeps();
  createWindow(CONFIG, off.deps);
  assert.notStrictEqual(off.created[0].webContents.devToolsOpen, true);

  const on = fakeDeps();
  createWindow({ ...CONFIG, webDebug: true }, on.deps);
  assert.strictEqual(on.created[0].webContents.devToolsOpen, true);
});

test("closing hides the window when closeToTray is on", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, true);
  assert.strictEqual(created[0].hidden, true);
});

test("closing really closes when closeToTray is off", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, closeToTray: false }, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, false);
  assert.strictEqual(created[0].hidden, false);
});

test("closeToTray is ignored when there is no tray to close to", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, trayIconEnabled: false }, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, false, "hiding with no tray would trap the window");
});

test("third-party links open in the system browser, Outlook links stay in-app", () => {
  const { created, deps, opened } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers.windowOpen;

  assert.deepStrictEqual(handler({ url: "https://example.com/report" }), {
    action: "deny",
  });
  assert.deepStrictEqual(opened, ["https://example.com/report"]);

  assert.deepStrictEqual(
    handler({ url: "https://outlook.office.com/mail/inbox/id/xyz" }),
    { action: "allow" },
  );
});

test("Microsoft sign-in URLs stay in-app so authentication can complete", () => {
  assert.strictEqual(
    shouldOpenExternally(
      "https://login.microsoftonline.com/common/oauth2/authorize",
      "https://outlook.office.com",
    ),
    false,
  );
  assert.strictEqual(
    shouldOpenExternally(
      "https://outlook.office365.com/mail",
      "https://outlook.office.com",
    ),
    false,
  );
});

test("non-http schemes are handed to the OS", () => {
  assert.strictEqual(
    shouldOpenExternally("mailto:someone@example.com", "https://outlook.office.com"),
    true,
  );
});

test("an unparseable URL is not opened externally", () => {
  assert.strictEqual(
    shouldOpenExternally("not a url", "https://outlook.office.com"),
    false,
  );
  assert.strictEqual(
    shouldOpenExternally(undefined, "https://outlook.office.com"),
    false,
  );
});
