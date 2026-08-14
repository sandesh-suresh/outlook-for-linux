const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const argv = require("../../app/config/index");

function withConfigDir(contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-config-"));
  if (contents !== undefined) {
    fs.writeFileSync(path.join(dir, "config.json"), contents);
  }
  return dir;
}

test("uses defaults when no config file exists", () => {
  const config = argv(withConfigDir(), "0.1.0", []);
  assert.strictEqual(config.url, "https://outlook.office.com");
  assert.strictEqual(config.trayIconEnabled, true);
  assert.strictEqual(config.isConfigFile, false);
});

test("the config file overrides defaults", () => {
  const dir = withConfigDir(
    JSON.stringify({ url: "https://outlook.live.com", trayIconEnabled: false }),
  );
  const config = argv(dir, "0.1.0", []);
  assert.strictEqual(config.url, "https://outlook.live.com");
  assert.strictEqual(config.trayIconEnabled, false);
  assert.strictEqual(config.isConfigFile, true);
});

test("CLI flags override the config file", () => {
  const dir = withConfigDir(
    JSON.stringify({ url: "https://outlook.live.com" }),
  );
  const config = argv(dir, "0.1.0", ["--url=https://outlook.office365.com"]);
  assert.strictEqual(config.url, "https://outlook.office365.com");
});

test("CLI flags override defaults for booleans", () => {
  const config = argv(withConfigDir(), "0.1.0", ["--trayIconEnabled=false"]);
  assert.strictEqual(config.trayIconEnabled, false);
});

test("a malformed config file is reported and does not throw", () => {
  const dir = withConfigDir("{ this is not json");
  const config = argv(dir, "0.1.0", []);
  assert.ok(config.error, "expected an error to be reported");
  // Defaults still apply so the app can start.
  assert.strictEqual(config.url, "https://outlook.office.com");
});

test("the returned config is deeply frozen", () => {
  const config = argv(withConfigDir(), "0.1.0", []);
  assert.strictEqual(Object.isFrozen(config), true);
  assert.strictEqual(Object.isFrozen(config.logConfig), true);
  assert.throws(() => {
    // Test files are CommonJS (sloppy mode), where writing to a frozen
    // property fails silently. The directive makes the failure observable.
    "use strict";
    config.url = "https://evil.example.com";
  }, TypeError);
});

test("the config records where it was loaded from", () => {
  const dir = withConfigDir();
  const config = argv(dir, "0.1.0", []);
  assert.strictEqual(config.configPath, dir);
});
