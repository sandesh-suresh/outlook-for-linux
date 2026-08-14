const test = require("node:test");
const assert = require("node:assert");
const {
  installNotificationShim,
} = require("../../app/browser/tools/notificationShim");

function harness() {
  const invoked = [];
  const channels = {};
  const ipcRenderer = {
    invoke: async (channel, payload) => {
      invoked.push({ channel, payload });
    },
    on: (channel, handler) => {
      channels[channel] = handler;
    },
  };
  const target = {};
  installNotificationShim(target, ipcRenderer);
  return { target, invoked, channels };
}

test("installs a Notification constructor on the target", () => {
  const { target } = harness();
  assert.strictEqual(typeof target.Notification, "function");
});

test("reports permission as granted without prompting", async () => {
  const { target } = harness();
  assert.strictEqual(target.Notification.permission, "granted");
  assert.strictEqual(await target.Notification.requestPermission(), "granted");
});

test("constructing a notification forwards it to the main process", async () => {
  const { target, invoked } = harness();
  new target.Notification("New mail", {
    body: "Quarterly report",
    icon: "data:image/png;base64,AAAA",
  });
  await Promise.resolve();

  assert.strictEqual(invoked.length, 1);
  assert.strictEqual(invoked[0].channel, "show-notification");
  assert.strictEqual(invoked[0].payload.title, "New mail");
  assert.strictEqual(invoked[0].payload.body, "Quarterly report");
  assert.strictEqual(invoked[0].payload.icon, "data:image/png;base64,AAAA");
  assert.strictEqual(typeof invoked[0].payload.id, "number");
});

test("each notification gets a distinct id", async () => {
  const { target, invoked } = harness();
  new target.Notification("a");
  new target.Notification("b");
  await Promise.resolve();
  assert.notStrictEqual(invoked[0].payload.id, invoked[1].payload.id);
});

test("tolerates being constructed with no options", async () => {
  const { target, invoked } = harness();
  assert.doesNotThrow(() => new target.Notification("New mail"));
  await Promise.resolve();
  assert.strictEqual(invoked[0].payload.body, "");
});

test("notification-closed fires onclose on the matching instance only", async () => {
  const { target, invoked, channels } = harness();
  const first = new target.Notification("a");
  const second = new target.Notification("b");
  await Promise.resolve();

  let firstClosed = 0;
  let secondClosed = 0;
  first.onclose = () => firstClosed++;
  second.onclose = () => secondClosed++;

  channels["notification-closed"]({}, { id: invoked[0].payload.id });

  assert.strictEqual(firstClosed, 1);
  assert.strictEqual(secondClosed, 0);
});

test("addEventListener('close') receives the close event", async () => {
  const { target, invoked, channels } = harness();
  const notification = new target.Notification("a");
  await Promise.resolve();

  let closed = 0;
  notification.addEventListener("close", () => closed++);
  channels["notification-closed"]({}, { id: invoked[0].payload.id });

  assert.strictEqual(closed, 1);
});

test("a close for an unknown id is ignored", () => {
  const { channels } = harness();
  assert.doesNotThrow(() =>
    channels["notification-closed"]({}, { id: 99999 }),
  );
  assert.doesNotThrow(() => channels["notification-closed"]({}, undefined));
});

test("a closed notification is not notified twice", async () => {
  const { target, invoked, channels } = harness();
  const notification = new target.Notification("a");
  await Promise.resolve();

  let closed = 0;
  notification.onclose = () => closed++;
  channels["notification-closed"]({}, { id: invoked[0].payload.id });
  channels["notification-closed"]({}, { id: invoked[0].payload.id });

  assert.strictEqual(closed, 1);
});

test("an IPC failure does not propagate to the page", async () => {
  const target = {};
  installNotificationShim(target, {
    invoke: async () => {
      throw new Error("no handler registered");
    },
    on: () => {},
  });

  assert.doesNotThrow(() => new target.Notification("New mail"));
  // Let the rejected invoke settle; an unhandled rejection would fail the run.
  await new Promise((resolve) => setImmediate(resolve));
});
