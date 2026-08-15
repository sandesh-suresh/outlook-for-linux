const test = require("node:test");
const assert = require("node:assert");
const unreadListObserver = require("../../app/browser/tools/unreadListObserver");

test("matches an unread message option's accessible name", () => {
  assert.strictEqual(
    unreadListObserver.isUnreadOptionLabel("Unread Jane Doe Test subject 9:15 AM preview text"),
    true,
  );
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel("Unread"), true);
});

test("does not match a read message or unrelated label", () => {
  assert.strictEqual(
    unreadListObserver.isUnreadOptionLabel("Jane Doe Test subject 9:15 AM preview text"),
    false,
  );
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel("Inbox"), false);
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel("Unreadable"), false);
});

test("rejects unusable input", () => {
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel(undefined), false);
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel(null), false);
  assert.strictEqual(unreadListObserver.isUnreadOptionLabel(12345), false);
});

test("counts only the labels that look unread", () => {
  assert.strictEqual(
    unreadListObserver.countUnreadLabels([
      "Unread Jane Doe Test subject A",
      "Jane Doe Test subject B",
      "Unread Jane Doe Test subject C",
      "Inbox",
    ]),
    2,
  );
  assert.strictEqual(unreadListObserver.countUnreadLabels([]), 0);
});

test("init is a no-op when the unread list logic is disabled", () => {
  // No DOM exists in this test process, so a no-op is observable as
  // "did not throw".
  assert.doesNotThrow(() =>
    unreadListObserver.init({ useUnreadListLogic: false }),
  );
});
