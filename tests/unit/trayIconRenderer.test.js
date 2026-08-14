const test = require("node:test");
const assert = require("node:assert");
const renderer = require("../../app/browser/tools/trayIconRenderer");

const FAKE_ICON = "data:image/png;base64,AAAA";

function harness(config = {}) {
  const sent = [];
  const invoked = [];
  renderer.ipcRenderer = {
    send: (channel, payload) => sent.push({ channel, payload }),
    invoke: async (channel, payload) => {
      invoked.push({ channel, payload });
    },
  };
  renderer.config = {
    disableBadgeCount: false,
    disableNotificationWindowFlash: false,
    ...config,
  };
  renderer.render = async () => FAKE_ICON;
  renderer.resetForTests();
  return { sent, invoked };
}

test("sends a rendered icon and the count for a non-zero unread count", async () => {
  const { sent, invoked } = harness();
  await renderer.updateActivityCount({ detail: { number: 3 } });

  assert.deepStrictEqual(sent, [
    {
      channel: "tray-update",
      payload: { icon: FAKE_ICON, flash: true, count: 3 },
    },
  ]);
  assert.deepStrictEqual(invoked, [{ channel: "set-badge-count", payload: 3 }]);
});

test("clearing to zero sends a null icon so main restores the base icon", async () => {
  const { sent } = harness();
  await renderer.updateActivityCount({ detail: { number: 2 } });
  await renderer.updateActivityCount({ detail: { number: 0 } });

  assert.strictEqual(sent.length, 2);
  assert.deepStrictEqual(sent[1].payload, {
    icon: null,
    flash: false,
    count: 0,
  });
});

test("repeating the same count sends nothing", async () => {
  const { sent } = harness();
  await renderer.updateActivityCount({ detail: { number: 5 } });
  await renderer.updateActivityCount({ detail: { number: 5 } });
  assert.strictEqual(sent.length, 1);
});

test("does not flash the window when flashing is disabled", async () => {
  const { sent } = harness({ disableNotificationWindowFlash: true });
  await renderer.updateActivityCount({ detail: { number: 1 } });
  assert.strictEqual(sent[0].payload.flash, false);
});

test("does not set the OS badge when badge counts are disabled", async () => {
  const { sent, invoked } = harness({ disableBadgeCount: true });
  await renderer.updateActivityCount({ detail: { number: 4 } });
  assert.strictEqual(sent.length, 1);
  assert.strictEqual(invoked.length, 0);
});

test("a render failure is swallowed and allows a retry of the same count", async () => {
  const { sent } = harness();
  renderer.render = async () => {
    throw new Error("canvas unavailable");
  };
  await renderer.updateActivityCount({ detail: { number: 7 } });
  assert.strictEqual(sent.length, 0);

  // The same count must not be deduplicated away after a failure.
  renderer.render = async () => FAKE_ICON;
  await renderer.updateActivityCount({ detail: { number: 7 } });
  assert.strictEqual(sent.length, 1);
});

test("a superseded render never overwrites a newer count", async () => {
  const { sent } = harness();
  let releaseSlowRender;
  renderer.render = (count) => {
    if (count === 1) {
      return new Promise((resolve) => {
        releaseSlowRender = () => resolve(FAKE_ICON);
      });
    }
    return Promise.resolve(FAKE_ICON);
  };

  const slow = renderer.updateActivityCount({ detail: { number: 1 } });
  await renderer.updateActivityCount({ detail: { number: 9 } });
  releaseSlowRender();
  await slow;

  assert.strictEqual(sent.length, 1);
  assert.strictEqual(sent[0].payload.count, 9);
});
