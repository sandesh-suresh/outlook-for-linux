const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const repoRoot = path.join(__dirname, "..", "..");
const pkg = require(path.join(repoRoot, "package.json"));

test("package.json declares the app entry point and identity", () => {
  assert.strictEqual(pkg.main, "app/index.js");
  assert.strictEqual(pkg.name, "outlook-for-linux");
  assert.strictEqual(pkg.license, "GPL-3.0-or-later");
});

test("package.json exposes the required scripts", () => {
  for (const script of ["lint", "test:unit", "start", "pack", "dist:linux"]) {
    assert.ok(pkg.scripts[script], `missing script: ${script}`);
  }
});

test("packaging targets AppImage, deb and rpm", () => {
  assert.deepStrictEqual([...pkg.build.linux.target].sort(), [
    "AppImage",
    "deb",
    "rpm",
  ]);
});

test("MVP-excluded dependencies are absent", () => {
  const deps = { ...pkg.dependencies, ...pkg.devDependencies };
  for (const forbidden of ["mqtt", "electron-updater", "electron-store"]) {
    assert.ok(!(forbidden in deps), `unexpected dependency: ${forbidden}`);
  }
});

test("LICENSE is GPL version 3", () => {
  const license = fs.readFileSync(path.join(repoRoot, "LICENSE"), "utf8");
  assert.match(license, /GNU GENERAL PUBLIC LICENSE/);
  assert.match(license, /Version 3/);
});
