const test = require("node:test");
const assert = require("node:assert");
const {
  validateIpcChannel,
  allowedChannels,
} = require("../../app/security/ipcValidator");

test("allows exactly the MVP channels", () => {
  assert.deepStrictEqual([...allowedChannels].sort(), [
    "get-config",
    "notification-closed",
    "set-badge-count",
    "show-notification",
    "tray-update",
    "unhandled-rejection",
    "window-error",
  ]);
});

test("permits an allowlisted channel", () => {
  assert.strictEqual(
    validateIpcChannel("show-notification", { title: "hi" }),
    true,
  );
});

test("blocks an unknown channel", () => {
  assert.strictEqual(validateIpcChannel("rm-rf-slash"), false);
});

test("blocks Teams-era channels that were not ported", () => {
  for (const channel of [
    "choose-desktop-media",
    "get-teams-settings",
    "graph-api-get-mail-messages",
  ]) {
    assert.strictEqual(
      validateIpcChannel(channel),
      false,
      `${channel} should be blocked`,
    );
  }
});

test("strips prototype-pollution keys from the payload", () => {
  const payload = JSON.parse('{"title":"hi","__proto__":{"polluted":true}}');
  assert.strictEqual(validateIpcChannel("show-notification", payload), true);
  assert.strictEqual({}.polluted, undefined);
  assert.strictEqual(Object.hasOwn(payload, "__proto__"), false);
});

test("strips pollution keys from nested payloads", () => {
  const payload = JSON.parse('{"outer":{"inner":{"__proto__":{"bad":1}}}}');
  validateIpcChannel("tray-update", payload);
  assert.strictEqual(Object.hasOwn(payload.outer.inner, "__proto__"), false);
});

test("tolerates a missing payload and cyclic payloads", () => {
  assert.strictEqual(validateIpcChannel("get-config"), true);
  const cyclic = { a: 1 };
  cyclic.self = cyclic;
  assert.strictEqual(validateIpcChannel("get-config", cyclic), true);
});
