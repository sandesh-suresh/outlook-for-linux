const test = require("node:test");
const assert = require("node:assert");
const mutationTitle = require("../../app/browser/tools/mutationTitle");

test("extracts the leading unread marker", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("(3) Inbox - Grace Hopper - Outlook"),
    3,
  );
  assert.strictEqual(
    mutationTitle.parseUnreadCount("(147) Mail - Outlook"),
    147,
  );
});

test("reports zero when there is no marker", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("Inbox - Grace Hopper - Outlook"),
    0,
  );
});

test("ignores a parenthesised value that is not a leading count", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("Mail (draft) - Outlook"),
    0,
  );
  assert.strictEqual(
    mutationTitle.parseUnreadCount("Re: (3) budget - Outlook"),
    0,
  );
});

test("rejects unusable input", () => {
  assert.strictEqual(mutationTitle.parseUnreadCount(undefined), null);
  assert.strictEqual(mutationTitle.parseUnreadCount(null), null);
  assert.strictEqual(mutationTitle.parseUnreadCount(12345), null);
});

test("rejects implausible counts", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("(99999) Mail - Outlook"),
    null,
  );
});

test("init is a no-op when the mutation title logic is disabled", () => {
  // No DOM exists in this test process, so a no-op is observable as
  // "did not throw".
  assert.doesNotThrow(() =>
    mutationTitle.init({ useMutationTitleLogic: false }),
  );
});
