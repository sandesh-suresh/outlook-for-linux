const test = require("node:test");
const assert = require("node:assert");
const { isNetworkError } = require("../../app/utils/networkErrors");

test("recognises Chromium network error codes", () => {
  assert.strictEqual(
    isNetworkError("Failed to load: ERR_INTERNET_DISCONNECTED"),
    true,
  );
  assert.strictEqual(isNetworkError("ERR_NAME_NOT_RESOLVED"), true);
});

test("treats destroyed-object errors as transient", () => {
  assert.strictEqual(isNetworkError("Object has been destroyed"), true);
});

test("treats failed script execution as a network symptom", () => {
  assert.strictEqual(isNetworkError("Script failed to execute"), true);
});

test("does not swallow genuine programming errors", () => {
  assert.strictEqual(isNetworkError("undefined is not a function"), false);
  assert.strictEqual(isNetworkError("ENOENT: no such file or directory"), false);
});

test("returns false for non-string input", () => {
  assert.strictEqual(isNetworkError(undefined), false);
  assert.strictEqual(isNetworkError(null), false);
  assert.strictEqual(isNetworkError(new Error("ERR_CONNECTION_RESET")), false);
});
