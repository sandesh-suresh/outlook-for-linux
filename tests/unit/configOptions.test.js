const test = require("node:test");
const assert = require("node:assert");
const options = require("../../app/config/options");
const { NETWORK_ERROR_PATTERNS } = require("../../app/config/defaults");

test("every option declares a default, a description and a type", () => {
  for (const [name, option] of Object.entries(options)) {
    assert.ok("default" in option, `${name} has no default`);
    assert.ok(option.describe, `${name} has no describe`);
    assert.ok(option.type, `${name} has no type`);
  }
});

test("the Outlook URL defaults to the work/school endpoint", () => {
  assert.strictEqual(options.url.default, "https://outlook.office.com");
  assert.strictEqual(options.url.type, "string");
});

test("tray and notifications are on by default", () => {
  assert.strictEqual(options.trayIconEnabled.default, true);
  assert.strictEqual(options.disableNotifications.default, false);
  assert.strictEqual(options.useMutationTitleLogic.default, true);
});

test("appIconType is constrained to the shipped icon variants", () => {
  assert.deepStrictEqual(options.appIconType.choices, [
    "default",
    "light",
    "dark",
  ]);
  assert.ok(options.appIconType.choices.includes(options.appIconType.default));
});

test("options.js is loadable without Electron", () => {
  // A require cycle through Electron would have thrown above; assert the
  // module stayed pure data.
  assert.strictEqual(typeof options, "object");
  assert.ok(!Array.isArray(options));
});

test("no Teams-specific or out-of-scope options leaked in", () => {
  for (const name of Object.keys(options)) {
    assert.ok(
      !/teams|mqtt|graph|meetup|spellcheck/i.test(name),
      `unexpected option: ${name}`,
    );
  }
});

test("network error patterns cover the transient connection failures", () => {
  for (const pattern of [
    "ERR_INTERNET_DISCONNECTED",
    "ERR_NAME_NOT_RESOLVED",
    "ERR_CONNECTION_RESET",
    "ERR_NETWORK_CHANGED",
  ]) {
    assert.ok(NETWORK_ERROR_PATTERNS.includes(pattern), `missing ${pattern}`);
  }
});
