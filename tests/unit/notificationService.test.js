const test = require("node:test");
const assert = require("node:assert");
const NotificationService = require("../../app/notifications/service");

function fakeDeps() {
  const created = [];
  class FakeNotification {
    constructor(options) {
      this.options = options;
      this.shown = false;
      this.listeners = {};
      created.push(this);
    }
    on(event, handler) {
      this.listeners[event] = handler;
      return this;
    }
    show() {
      this.shown = true;
    }
  }
  return {
    created,
    deps: {
      Notification: FakeNotification,
      nativeImage: {
        createFromDataURL: (url) => ({
          from: "dataURL",
          url,
          isEmpty: () => false,
        }),
      },
    },
  };
}

function fakeWindow() {
  const sent = [];
  return {
    sent,
    shown: false,
    focused: false,
    isVisible: () => false,
    show() {
      this.shown = true;
    },
    focus() {
      this.focused = true;
    },
    webContents: {
      send: (channel, payload) => sent.push({ channel, payload }),
    },
  };
}

const CONFIG = {
  disableNotifications: false,
  defaultNotificationUrgency: "normal",
};

test("shows a native notification with the supplied title and body", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);
  await service.show({ title: "New mail", body: "Quarterly report" });

  assert.strictEqual(created.length, 1);
  assert.strictEqual(created[0].options.title, "New mail");
  assert.strictEqual(created[0].options.body, "Quarterly report");
  assert.strictEqual(created[0].options.urgency, "normal");
  assert.strictEqual(created[0].shown, true);
});

test("honours the configured default urgency and per-notification override", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(
    { ...CONFIG, defaultNotificationUrgency: "low" },
    deps,
  );
  await service.show({ title: "a" });
  await service.show({ title: "b", urgency: "critical" });

  assert.strictEqual(created[0].options.urgency, "low");
  assert.strictEqual(created[1].options.urgency, "critical");
});

test("shows nothing when notifications are disabled", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(
    { ...CONFIG, disableNotifications: true },
    deps,
  );
  await service.show({ title: "New mail" });
  assert.strictEqual(created.length, 0);
});

test("accepts a data URL icon and ignores any other icon source", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);

  await service.show({ title: "a", icon: "data:image/png;base64,AAAA" });
  assert.strictEqual(created[0].options.icon.from, "dataURL");
  assert.strictEqual(
    created[0].options.icon.url,
    "data:image/png;base64,AAAA",
  );

  await service.show({
    title: "b",
    icon: "https://outlook.office.com/avatar.png",
  });
  assert.strictEqual(created[1].options.icon, undefined);
});

test("clicking a notification shows and focuses the window", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);
  const window = fakeWindow();
  service.setWindow(window);

  await service.show({ id: 7, title: "New mail" });
  created[0].listeners.click();

  assert.strictEqual(window.shown, true);
  assert.strictEqual(window.focused, true);
  assert.deepStrictEqual(window.sent, [
    { channel: "notification-closed", payload: { id: 7 } },
  ]);
});

test("dismissing a notification notifies the renderer once", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);
  const window = fakeWindow();
  service.setWindow(window);

  await service.show({ id: 9, title: "New mail" });
  created[0].listeners.close();
  created[0].listeners.close();

  assert.strictEqual(window.sent.length, 1);
  assert.deepStrictEqual(window.sent[0].payload, { id: 9 });
});

test("a missing window does not break the click handler", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);
  await service.show({ id: 1, title: "New mail" });
  assert.doesNotThrow(() => created[0].listeners.click());
});

test("malformed options are tolerated", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);

  await service.show(undefined);
  await service.show({});

  assert.strictEqual(created.length, 2);
  assert.strictEqual(created[1].options.title, "Microsoft Outlook");
  assert.strictEqual(created[1].options.body, "");
});

test("initialize registers the show-notification handler", async () => {
  const { created, deps } = fakeDeps();
  const handlers = {};
  const service = new NotificationService(CONFIG, deps);
  service.initialize({
    handle: (channel, handler) => (handlers[channel] = handler),
  });

  await handlers["show-notification"]({}, { title: "New mail" });
  assert.strictEqual(created.length, 1);
});

test("a notification failure is swallowed", async () => {
  const { deps } = fakeDeps();
  deps.Notification = class {
    show() {
      throw new Error("no notification daemon");
    }
    on() {
      return this;
    }
  };
  const service = new NotificationService(CONFIG, deps);
  await assert.doesNotReject(() => service.show({ title: "New mail" }));
});
