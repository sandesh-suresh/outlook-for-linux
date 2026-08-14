const test = require("node:test");
const assert = require("node:assert");
const logger = require("../../app/config/logger");

function fakeLog() {
  return {
    hooks: [],
    functions: { log() {}, info() {}, debug() {}, warn() {}, error() {} },
    transports: { console: { level: "info" }, file: { level: false } },
    initialized: false,
    initialize() {
      this.initialized = true;
    },
  };
}

test("installs a hook that strips PII from log data", () => {
  const log = fakeLog();
  logger.init({ transports: { console: { level: "debug" } } }, log);

  assert.strictEqual(log.initialized, true);
  assert.strictEqual(log.hooks.length, 1);

  const message = { data: ["contact frank@contoso.com", { token: "abc" }] };
  const hooked = log.hooks[0](message);
  assert.deepStrictEqual(hooked.data, [
    "contact [EMAIL]",
    { token: "[REDACTED]" },
  ]);
});

test("merges the supplied transport config into the log instance", () => {
  const log = fakeLog();
  logger.init({ transports: { console: { level: "debug" } } }, log);
  assert.strictEqual(log.transports.console.level, "debug");
  // Untouched keys survive the merge.
  assert.strictEqual(log.transports.file.level, false);
});

test('the "console" config leaves logging untouched', () => {
  const log = fakeLog();
  logger.init("console", log);
  assert.strictEqual(log.initialized, false);
  assert.strictEqual(log.hooks.length, 0);
});

test("a falsy config silences the console", () => {
  const original = { ...console };
  try {
    logger.init(false, fakeLog());
    // No-op functions return undefined and must not throw.
    assert.strictEqual(console.info("swallowed"), undefined);
    assert.strictEqual(console.error("swallowed"), undefined);
  } finally {
    Object.assign(console, original);
  }
});

test("merging refuses prototype-polluting keys", () => {
  const log = fakeLog();
  const malicious = JSON.parse('{"__proto__": {"polluted": true}}');
  logger.init(malicious, log);
  assert.strictEqual({}.polluted, undefined);
});
