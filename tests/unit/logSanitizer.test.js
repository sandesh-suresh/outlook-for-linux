const test = require("node:test");
const assert = require("node:assert");
const {
  sanitize,
  sanitizeObject,
  sanitizeLogData,
  containsPII,
} = require("../../app/utils/logSanitizer");

test("redacts email addresses", () => {
  assert.strictEqual(
    sanitize("login failed for alice.smith@contoso.com"),
    "login failed for [EMAIL]",
  );
});

test("redacts bearer tokens and access tokens", () => {
  assert.strictEqual(sanitize("Bearer abc123.def-456"), "Bearer [TOKEN]");
  assert.strictEqual(sanitize("access_token=xyz789"), "access_token=[REDACTED]");
});

test("redacts URL query parameters but keeps the path", () => {
  assert.strictEqual(
    sanitize("GET https://outlook.office.com/mail?token=secret&id=42"),
    "GET https://outlook.office.com/mail?[PARAMS]",
  );
});

test("redacts the user's home directory", () => {
  assert.strictEqual(
    sanitize("reading /home/alice/.config/outlook-for-linux/config.json"),
    "reading /home/[USER]/.config/outlook-for-linux/config.json",
  );
});

test("truncates UUIDs to their first segment", () => {
  assert.strictEqual(
    sanitize("id 123e4567-e89b-12d3-a456-426614174000"),
    "id 123e4567...",
  );
});

test("handles null and undefined without throwing", () => {
  assert.strictEqual(sanitize(null), "null");
  assert.strictEqual(sanitize(undefined), "undefined");
});

test("sanitizeObject redacts sensitive keys wholesale", () => {
  const result = sanitizeObject({
    user: "bob@contoso.com",
    accessToken: "should-not-appear",
    nested: { password: "hunter2", note: "ok" },
  });
  assert.strictEqual(result.user, "[EMAIL]");
  assert.strictEqual(result.accessToken, "[REDACTED]");
  assert.strictEqual(result.nested.password, "[REDACTED]");
  assert.strictEqual(result.nested.note, "ok");
});

test("sanitizeObject survives circular references", () => {
  const obj = { name: "root" };
  obj.self = obj;
  assert.strictEqual(sanitizeObject(obj).self, "[Circular]");
});

test("sanitizeObject preserves Error shape while redacting the message", () => {
  const result = sanitizeObject(new Error("failed for carol@contoso.com"));
  assert.ok(result instanceof Error);
  assert.strictEqual(result.message, "failed for [EMAIL]");
});

test("sanitizeLogData maps over mixed log arguments", () => {
  const result = sanitizeLogData(["user dave@contoso.com", { token: "t" }, 42]);
  assert.deepStrictEqual(result, ["user [EMAIL]", { token: "[REDACTED]" }, 42]);
});

test("containsPII detects and clears regex state between calls", () => {
  assert.strictEqual(containsPII("erin@contoso.com"), true);
  assert.strictEqual(containsPII("erin@contoso.com"), true);
  assert.strictEqual(containsPII("nothing to see"), false);
});
