const test = require("node:test");
const assert = require("node:assert");
const toastObserver = require("../../app/browser/tools/toastObserver");

test("uses the first text run as the title and the second as the body", () => {
  assert.deepStrictEqual(
    toastObserver.buildToastMessage(["Jane Doe", "Quarterly report"], "Outlook"),
    { title: "Jane Doe", body: "Quarterly report" },
  );
});

test("puts three or more runs on their own line in the body", () => {
  assert.deepStrictEqual(
    toastObserver.buildToastMessage(
      ["Jane Doe", "Quarterly report", "Here is the summary"],
      "Outlook",
    ),
    { title: "Jane Doe", body: "Quarterly report\nHere is the summary" },
  );
});

test("trims and drops empty runs before building the message", () => {
  assert.deepStrictEqual(
    toastObserver.buildToastMessage(
      ["  Jane Doe  ", "   ", "Quarterly report  "],
      "Outlook",
    ),
    { title: "Jane Doe", body: "Quarterly report" },
  );
});

test("falls back to the app title when only one run is found", () => {
  assert.deepStrictEqual(
    toastObserver.buildToastMessage(["You have a new message"], "Outlook"),
    { title: "Outlook", body: "You have a new message" },
  );
});

test("drops decorative runs shorter than 4 characters before picking the title", () => {
  assert.deepStrictEqual(
    toastObserver.buildToastMessage(["Al", "Jane Doe", "Quarterly report"], "Outlook"),
    { title: "Jane Doe", body: "Quarterly report" },
  );
});

test("treats an all-decorative run set as unusable", () => {
  assert.strictEqual(toastObserver.buildToastMessage(["Al", "Hi"], "Outlook"), null);
});

test("returns null when there are no usable runs", () => {
  assert.strictEqual(toastObserver.buildToastMessage([], "Outlook"), null);
  assert.strictEqual(toastObserver.buildToastMessage(["   ", ""], "Outlook"), null);
  assert.strictEqual(toastObserver.buildToastMessage(undefined, "Outlook"), null);
});

test("init does not throw with no DOM", () => {
  // No DOM exists in this test process, so a no-op is observable as
  // "did not throw".
  assert.doesNotThrow(() => toastObserver.init({ appTitle: "Outlook" }));
});
