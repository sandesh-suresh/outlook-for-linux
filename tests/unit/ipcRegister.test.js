const test = require("node:test");
const assert = require("node:assert");
const {
  createGuardedIpcMain,
  registerIpcHandlers,
} = require("../../app/ipc/register");

function fakeIpcMain() {
  const on = {};
  const handle = {};
  return {
    on: (channel, handler) => (on[channel] = handler),
    handle: (channel, handler) => (handle[channel] = handler),
    registered: { on, handle },
  };
}

test("registering a channel that is not allowlisted throws at startup", () => {
  const guarded = createGuardedIpcMain(fakeIpcMain());
  assert.throws(
    () => guarded.handle("definitely-not-allowed", () => {}),
    /not allowlisted/i,
  );
});

test("an allowlisted channel reaches the real ipcMain", () => {
  const ipcMain = fakeIpcMain();
  const guarded = createGuardedIpcMain(ipcMain);
  guarded.handle("get-config", () => "ok");
  guarded.on("tray-update", () => {});

  assert.strictEqual(
    typeof ipcMain.registered.handle["get-config"],
    "function",
  );
  assert.strictEqual(typeof ipcMain.registered.on["tray-update"], "function");
});

test("payloads are sanitized before the handler sees them", () => {
  const ipcMain = fakeIpcMain();
  const guarded = createGuardedIpcMain(ipcMain);

  let seen;
  guarded.on("tray-update", (_event, payload) => (seen = payload));

  const payload = JSON.parse('{"count":3,"__proto__":{"polluted":true}}');
  ipcMain.registered.on["tray-update"]({}, payload);

  assert.strictEqual(seen.count, 3);
  assert.strictEqual({}.polluted, undefined);
});

test("get-config returns the resolved config", async () => {
  const ipcMain = fakeIpcMain();
  const config = { url: "https://outlook.office.com", trayIconEnabled: true };
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config,
    setBadgeCount: () => {},
  });

  assert.deepStrictEqual(
    await ipcMain.registered.handle["get-config"]({}),
    config,
  );
});

test("set-badge-count forwards a valid count and rejects nonsense", async () => {
  const ipcMain = fakeIpcMain();
  const counts = [];
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: (n) => counts.push(n),
  });

  const handler = ipcMain.registered.handle["set-badge-count"];
  await handler({}, 4);
  await handler({}, 0);
  await handler({}, "many");
  await handler({}, -1);

  assert.deepStrictEqual(counts, [4, 0]);
});

test("a failing setBadgeCount does not reject the invoke", async () => {
  const ipcMain = fakeIpcMain();
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: () => {
      throw new Error("unsupported on this desktop");
    },
  });

  await assert.doesNotReject(() =>
    ipcMain.registered.handle["set-badge-count"]({}, 2),
  );
});

test("renderer error channels are registered and do not throw", () => {
  const ipcMain = fakeIpcMain();
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: () => {},
  });

  assert.doesNotThrow(() =>
    ipcMain.registered.on["window-error"]({}, { message: "boom", line: 12 }),
  );
  assert.doesNotThrow(() =>
    ipcMain.registered.on["unhandled-rejection"]({}, { reason: "boom" }),
  );
  assert.doesNotThrow(() =>
    ipcMain.registered.on["window-error"]({}, undefined),
  );
});
