const test = require("node:test");
const assert = require("node:assert");
const {
  createWindow,
  shouldOpenExternally,
  buildContextMenuTemplate,
} = require("../../app/mainAppWindow");

function fakeDeps() {
  const opened = [];
  const created = [];
  const poppedUp = [];
  class FakeMenu {
    static buildFromTemplate(template) {
      return {
        template,
        popup() {
          poppedUp.push(template);
        },
      };
    }
  }
  class FakeBrowserWindow {
    constructor(options) {
      this.options = options;
      this.loadedUrl = null;
      this.listeners = {};
      this.hidden = false;
      this.webContents = {
        handlers: {},
        userAgent:
          "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) outlook-for-linux/0.3.1 Chrome/120.0.0.0 Electron/28.0.0 Safari/537.36",
        setWindowOpenHandler(handler) {
          this.handlers.windowOpen = handler;
        },
        on(event, handler) {
          this.handlers[event] = handler;
        },
        openDevTools() {
          this.devToolsOpen = true;
        },
        closeDevTools() {
          this.devToolsOpen = false;
          this.closeDevToolsCalls = (this.closeDevToolsCalls || 0) + 1;
        },
        getUserAgent() {
          return this.userAgent;
        },
        setUserAgent(userAgent) {
          this.userAgent = userAgent;
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
    poppedUp,
    deps: {
      BrowserWindow: FakeBrowserWindow,
      shell: { openExternal: (url) => opened.push(url) },
      Menu: FakeMenu,
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

test("strips the app-name and Electron tokens from the UA so Outlook treats this as a supported Chrome", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);

  assert.strictEqual(
    created[0].webContents.userAgent,
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  );
});

test("blocks the DevTools keyboard shortcut when web debugging is off", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers["before-input-event"];

  const f12Event = { prevented: false, preventDefault() { this.prevented = true; } };
  handler(f12Event, { key: "F12", control: false, meta: false, shift: false });
  assert.strictEqual(f12Event.prevented, true);

  const ctrlShiftIEvent = { prevented: false, preventDefault() { this.prevented = true; } };
  handler(ctrlShiftIEvent, { key: "i", control: true, meta: false, shift: true });
  assert.strictEqual(ctrlShiftIEvent.prevented, true);

  const plainIEvent = { prevented: false, preventDefault() { this.prevented = true; } };
  handler(plainIEvent, { key: "i", control: false, meta: false, shift: false });
  assert.strictEqual(plainIEvent.prevented, false);
});

test("closes DevTools immediately if it opens by any other means, when web debugging is off", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers["devtools-opened"];

  handler();

  assert.strictEqual(created[0].webContents.closeDevToolsCalls, 1);
});

test("leaves DevTools reachable when web debugging is on", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, webDebug: true }, deps);

  assert.strictEqual(created[0].webContents.handlers["before-input-event"], undefined);
  assert.strictEqual(created[0].webContents.handlers["devtools-opened"], undefined);
});

test("a popped-out window also has DevTools blocked when web debugging is off", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);
  const didCreateWindow = created[0].webContents.handlers["did-create-window"];

  const child = new deps.BrowserWindow({});
  didCreateWindow(child, { url: "https://outlook.office.com/mail/deeplink" });

  assert.strictEqual(typeof child.webContents.handlers["before-input-event"], "function");
  assert.strictEqual(typeof child.webContents.handlers["devtools-opened"], "function");
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

test("about:blank stays in-app, since it's a popup bootstrapping before Outlook navigates it", () => {
  assert.strictEqual(
    shouldOpenExternally("about:blank", "https://outlook.office.com"),
    false,
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

test("an editable region gets the standard edit menu, respecting what's actually enabled", () => {
  const template = buildContextMenuTemplate({
    isEditable: true,
    selectionText: "",
    editFlags: {
      canUndo: true,
      canRedo: false,
      canCut: false,
      canCopy: false,
      canPaste: true,
      canSelectAll: true,
    },
  });

  const byRole = Object.fromEntries(
    template.filter((item) => item.role).map((item) => [item.role, item]),
  );
  assert.strictEqual(byRole.undo.enabled, true);
  assert.strictEqual(byRole.redo.enabled, false);
  assert.strictEqual(byRole.cut.enabled, false);
  assert.strictEqual(byRole.copy.enabled, false);
  assert.strictEqual(byRole.paste.enabled, true);
  assert.strictEqual(byRole.selectAll.enabled, true);
});

test("selected read-only text gets just a Copy entry", () => {
  const template = buildContextMenuTemplate({
    isEditable: false,
    selectionText: "some selected text",
    editFlags: { canCopy: true },
  });

  assert.deepStrictEqual(
    template.map((item) => item.role),
    ["copy"],
  );
  assert.strictEqual(template[0].enabled, true);
});

test("right-clicking non-interactive, non-selected content shows no menu", () => {
  const template = buildContextMenuTemplate({
    isEditable: false,
    selectionText: "",
    editFlags: {},
  });

  assert.deepStrictEqual(template, []);
});

test("right-clicking while editing pops up the edit menu", () => {
  const { created, deps, poppedUp } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers["context-menu"];

  handler({}, { isEditable: true, selectionText: "", editFlags: {} });

  assert.strictEqual(poppedUp.length, 1);
});

test("right-clicking empty, non-editable space pops up nothing", () => {
  const { created, deps, poppedUp } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers["context-menu"];

  handler({}, { isEditable: false, selectionText: "", editFlags: {} });

  assert.strictEqual(poppedUp.length, 0);
});

test("a popped-out window (e.g. 'Open in a new window') also gets a context menu", () => {
  const { created, deps, poppedUp } = fakeDeps();
  createWindow(CONFIG, deps);
  const didCreateWindow = created[0].webContents.handlers["did-create-window"];

  const child = new deps.BrowserWindow({});
  didCreateWindow(child, { url: "https://outlook.office.com/mail/deeplink" });
  const childHandler = child.webContents.handlers["context-menu"];
  childHandler({}, { isEditable: true, selectionText: "", editFlags: {} });

  assert.strictEqual(poppedUp.length, 1);
});
