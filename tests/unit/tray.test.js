const test = require("node:test");
const assert = require("node:assert");
const ApplicationTray = require("../../app/menus/tray");

function fakeDeps() {
  const created = [];
  class FakeTray {
    constructor(image) {
      this.image = image;
      this.tooltip = null;
      this.menu = null;
      this.destroyed = false;
      this.listeners = {};
      created.push(this);
    }
    setToolTip(text) {
      this.tooltip = text;
    }
    setContextMenu(menu) {
      this.menu = menu;
    }
    setImage(image) {
      this.image = image;
    }
    on(event, handler) {
      this.listeners[event] = handler;
    }
    destroy() {
      this.destroyed = true;
    }
  }
  return {
    created,
    deps: {
      Tray: FakeTray,
      Menu: { buildFromTemplate: (template) => ({ template }) },
      nativeImage: {
        createFromDataURL: (url) => ({ from: "dataURL", url }),
        createFromPath: (p) => ({ from: "path", path: p }),
      },
    },
  };
}

function fakeWindow() {
  return {
    visible: true,
    flashed: null,
    focused: false,
    isVisible() {
      return this.visible;
    },
    show() {
      this.visible = true;
    },
    hide() {
      this.visible = false;
    },
    focus() {
      this.focused = true;
    },
    flashFrame(value) {
      this.flashed = value;
    },
  };
}

function fakeIpcMain() {
  const handlers = {};
  return {
    handlers,
    on(channel, handler) {
      handlers[channel] = handler;
    },
  };
}

const CONFIG = { appTitle: "Microsoft Outlook", trayIconEnabled: true };

test("creates a tray with the base icon, tooltip and a context menu", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  assert.strictEqual(created.length, 1);
  assert.deepStrictEqual(created[0].image, {
    from: "path",
    path: "/icons/base.png",
  });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook");
  assert.ok(created[0].menu, "a context menu must be attached");
});

test("does not create a tray when the tray icon is disabled", () => {
  const { created, deps } = fakeDeps();
  const ipcMain = fakeIpcMain();
  const tray = new ApplicationTray(
    fakeWindow(),
    { ...CONFIG, trayIconEnabled: false },
    "/icons/base.png",
    deps,
  );
  tray.initialize(ipcMain);

  assert.strictEqual(created.length, 0);
  assert.strictEqual(ipcMain.handlers["tray-update"], undefined);
});

test("the context menu starts with an unread-count item reading 'No unread emails'", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  assert.strictEqual(created[0].menu.template[0].label, "No unread emails");
});

test("a tray-update with a count of one updates the unread-count item to the singular form", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  tray.updateTrayImage({ icon: null, flash: false, count: 1 });

  assert.strictEqual(created[0].menu.template[0].label, "1 unread email");
});

test("a tray-update with a count above one updates the unread-count item to the plural form", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  tray.updateTrayImage({ icon: null, flash: false, count: 4 });

  assert.strictEqual(created[0].menu.template[0].label, "4 unread emails");
});

test("a tray-update back to zero resets the unread-count item", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  tray.updateTrayImage({ icon: null, flash: false, count: 4 });
  tray.updateTrayImage({ icon: null, flash: false, count: 0 });

  assert.strictEqual(created[0].menu.template[0].label, "No unread emails");
});

test("clicking the unread-count item shows and focuses a hidden window", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  window.visible = false;
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  created[0].menu.template[0].click();

  assert.strictEqual(window.visible, true);
  assert.strictEqual(window.focused, true);
});

test("a tray-update with a count swaps the image and flashes the window", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  const ipcMain = fakeIpcMain();
  tray.initialize(ipcMain);

  ipcMain.handlers["tray-update"](
    {},
    {
      icon: "data:image/png;base64,AAAA",
      flash: true,
      count: 4,
    },
  );

  assert.deepStrictEqual(created[0].image, {
    from: "dataURL",
    url: "data:image/png;base64,AAAA",
  });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook (4)");
  assert.strictEqual(window.flashed, true);
});

test("a zero count restores the base icon and plain tooltip", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  tray.updateTrayImage({ icon: null, flash: false, count: 0 });

  assert.deepStrictEqual(created[0].image, {
    from: "path",
    path: "/icons/base.png",
  });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook");
  assert.strictEqual(window.flashed, false);
});

test("clicking the tray shows and focuses a hidden window", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  window.visible = false;
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  created[0].listeners.click();

  assert.strictEqual(window.visible, true);
  assert.strictEqual(window.focused, true);
});

test("a malformed tray-update payload does not throw", () => {
  const { deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  assert.doesNotThrow(() => tray.updateTrayImage(undefined));
  assert.doesNotThrow(() => tray.updateTrayImage({}));
});

test("close destroys the tray and is safe to call twice", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(
    fakeWindow(),
    CONFIG,
    "/icons/base.png",
    deps,
  );
  tray.initialize(fakeIpcMain());

  tray.close();
  tray.close();

  assert.strictEqual(created[0].destroyed, true);
});
