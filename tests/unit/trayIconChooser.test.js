const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const TrayIconChooser = require("../../app/browser/tools/trayIconChooser");

test("returns the bundled default icon when no custom icon is set", () => {
  const file = new TrayIconChooser({
    appIcon: "",
    appIconType: "default",
  }).getFile();
  assert.ok(path.isAbsolute(file));
  assert.strictEqual(path.basename(file), "icon-96x96.png");
  assert.ok(fs.existsSync(file), "the bundled icon must exist on disk");
});

test("selects the monochrome variants by icon type", () => {
  for (const [type, expected] of [
    ["light", "icon-monochrome-light-96x96.png"],
    ["dark", "icon-monochrome-dark-96x96.png"],
  ]) {
    const file = new TrayIconChooser({
      appIcon: "",
      appIconType: type,
    }).getFile();
    assert.strictEqual(path.basename(file), expected);
    assert.ok(fs.existsSync(file), `${expected} must exist on disk`);
  }
});

test("a custom icon path wins over the bundled icons", () => {
  const custom = "/opt/icons/my-outlook.png";
  const file = new TrayIconChooser({
    appIcon: custom,
    appIconType: "default",
  }).getFile();
  assert.strictEqual(file, custom);
});

test("a whitespace-only custom icon path falls back to the bundled icon", () => {
  const file = new TrayIconChooser({
    appIcon: "   ",
    appIconType: "default",
  }).getFile();
  assert.strictEqual(path.basename(file), "icon-96x96.png");
});

test("an unknown icon type falls back to the default icon", () => {
  const file = new TrayIconChooser({
    appIcon: "",
    appIconType: "chartreuse",
  }).getFile();
  assert.strictEqual(path.basename(file), "icon-96x96.png");
});
