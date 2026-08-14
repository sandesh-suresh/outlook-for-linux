const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { AppConfiguration } = require("../../app/appConfiguration");

function tempConfigDir(contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-appconfig-"));
  if (contents) {
    fs.writeFileSync(path.join(dir, "config.json"), JSON.stringify(contents));
  }
  return dir;
}

test("exposes the resolved startup config", () => {
  const dir = tempConfigDir({ url: "https://outlook.live.com" });
  const appConfig = new AppConfiguration(dir, "0.1.0", []);
  assert.strictEqual(appConfig.startupConfig.url, "https://outlook.live.com");
  assert.strictEqual(appConfig.configPath, dir);
});

test("the startup config cannot be replaced from outside", () => {
  const appConfig = new AppConfiguration(tempConfigDir(), "0.1.0", []);
  assert.throws(() => {
    // Test files are CommonJS (sloppy mode), where assigning to a getter-only
    // accessor fails silently. The directive makes the failure observable.
    "use strict";
    appConfig.startupConfig = { url: "https://evil.example.com" };
  }, TypeError);
});

test("private state is not enumerable on the instance", () => {
  const appConfig = new AppConfiguration(tempConfigDir(), "0.1.0", []);
  assert.deepStrictEqual(Object.keys(appConfig), []);
});
