const test = require("node:test");
const assert = require("node:assert");
const CommandLineManager = require("../../app/startup/commandLine");

function fakeApp() {
  const switches = [];
  return {
    switches,
    commandLine: {
      appendSwitch(name, value) {
        switches.push(value === undefined ? name : `${name}=${value}`);
      },
    },
  };
}

test("sets the WM_CLASS so Linux desktops group the window correctly", () => {
  const app = fakeApp();
  CommandLineManager.applySwitches({ disableGpu: false }, app);
  assert.ok(app.switches.includes("class=outlook-for-linux"));
});

test("disables GPU acceleration when configured", () => {
  const app = fakeApp();
  CommandLineManager.applySwitches({ disableGpu: true }, app);
  assert.ok(app.switches.includes("disable-gpu"));
  assert.ok(app.switches.includes("disable-gpu-compositing"));
  assert.ok(app.switches.includes("disable-software-rasterizer"));
});

test("leaves GPU acceleration alone by default", () => {
  const app = fakeApp();
  CommandLineManager.applySwitches({ disableGpu: false }, app);
  assert.ok(!app.switches.some((s) => s.startsWith("disable-gpu")));
});

test("tolerates a config with no GPU key", () => {
  const app = fakeApp();
  assert.doesNotThrow(() => CommandLineManager.applySwitches({}, app));
});
