# outlook-for-linux MVP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an Electron desktop wrapper for the Outlook web client on Linux that loads Outlook web in a native window, forwards Outlook's browser notifications to native OS notifications, shows a tray icon with a best-effort unread badge, reads layered configuration, and packages as AppImage/deb/rpm.

**Architecture:** A main Electron process (`app/index.js`) owns configuration, the single `BrowserWindow`, the tray, and native notifications. A preload script running inside the Outlook page replaces `globalThis.Notification` with a shim that forwards notification requests to main over IPC, and a `MutationObserver` on `<title>` extracts the unread count and drives the tray badge. Every IPC channel passes an allowlist validator, and all logging runs through a PII sanitizer. Conventions are ported from teams-for-linux so the two projects stay structurally recognisable.

**Tech Stack:** Electron 42, `electron-log` (structured logging), `yargs` (config/CLI parsing), Node's built-in `node --test` runner, ESLint 10 (flat config), electron-builder 26.

**Spec:** `docs/superpowers/specs/2026-08-14-outlook-for-linux-mvp-design.md`

## Global Constraints

Every task's requirements implicitly include this section.

- **License:** GPL-3.0-or-later (matching teams-for-linux).
- **Default Outlook URL:** `https://outlook.office.com`, overridable via config file and CLI.
- **Config file location:** `~/.config/outlook-for-linux/config.json` (i.e. Electron's `app.getPath("userData")` for an app named `outlook-for-linux`).
- **Config precedence:** defaults → config file → CLI flags. Config is treated as immutable after startup.
- **Packaging targets:** `AppImage`, `deb`, `rpm` via electron-builder.
- **Code style:** no `var` — `const` by default, `let` only for reassignment; `async`/`await` over promise chains; JavaScript `#property` private class fields; arrow functions for concise callbacks. CommonJS modules (`require`/`module.exports`) throughout `app/`.
- **Logging / PII:** never log email addresses, tokens, credentials, URL query parameters, or the user's home path. All `electron-log` output passes through the log sanitizer. Use `console.error` / `warn` / `info` / `debug` with bracketed subsystem prefixes, e.g. `console.info("[TRAY] Tray created")`.
- **DOM defensiveness:** every access to Outlook's DOM is wrapped in `try`/`catch` with a log-and-no-op fallback, never a throw. Outlook's DOM can change without notice.
- **Tests:** unit tests live in `tests/unit/*.test.js` and run under `node --test`. Electron-dependent collaborators are injected through constructors so units are testable without launching Electron.
- **Out of scope for the MVP** (do not add, even opportunistically): MQTT, Microsoft Graph, multi-account profiles/partitions, custom CSS/backgrounds, screen sharing, WebAuthn/Intune SSO, spellcheck, global shortcuts, quick chat, auto-updater, E2E tests.

## File Structure

| File | Responsibility |
| --- | --- |
| `package.json` | Metadata, scripts, dependencies, electron-builder `build` block. |
| `eslint.config.mjs` | Flat ESLint config enforcing `no-var` / `eqeqeq`. |
| `.gitignore`, `LICENSE`, `README.md` | Repo hygiene, GPL-3.0 text, user-facing intro. |
| `app/index.js` | Main entry: global error handlers, config load, window + tray + notification wiring, app lifecycle. |
| `app/config/defaults.js` | Pure data: `NETWORK_ERROR_PATTERNS` and shared default constants. No Electron imports. |
| `app/config/options.js` | Pure data: the yargs option schema (single source of truth for the config surface). No Electron imports. |
| `app/config/logger.js` | `electron-log` initialisation plus the PII sanitizer hook. |
| `app/config/index.js` | Loads config file, merges with defaults and CLI via yargs, returns the frozen config object. |
| `app/appConfiguration/index.js` | `AppConfiguration` class wrapping the resolved startup config behind `#private` fields. |
| `app/security/ipcValidator.js` | IPC channel allowlist + payload prototype-pollution sanitisation. |
| `app/utils/logSanitizer.js` | PII redaction for log strings/objects. |
| `app/utils/networkErrors.js` | `isNetworkError(message)` classifier shared by the global error handlers. |
| `app/startup/commandLine.js` | Chromium/Electron command-line switches derived from config, applied before the first window. |
| `app/ipc/register.js` | Guards every IPC registration through the allowlist; registers the main process's own channels. |
| `app/notifications/service.js` | `NotificationService`: `show-notification` IPC → native `Notification`. |
| `app/menus/tray.js` | `ApplicationTray`: creates the tray, handles `tray-update`, owns the context menu. |
| `app/mainAppWindow/index.js` | `createWindow(config)`: creates the `BrowserWindow`, loads `config.url`, wires the preload, routes external links. |
| `app/browser/preload.js` | Runs in the Outlook page: installs the notification shim, forwards renderer errors, inits browser tools. |
| `app/browser/tools/notificationShim.js` | Replaces the page's `Notification` constructor with an IPC-forwarding one. |
| `app/browser/tools/mutationTitle.js` | `<title>` observer → `unread-count` custom event. |
| `app/browser/tools/trayIconRenderer.js` | `unread-count` → badge-composited icon → `tray-update` IPC. |
| `app/browser/tools/trayIconChooser.js` | Resolves which tray icon file to use from config. |
| `app/assets/icons/*.png` | Runtime tray icons (default/light/dark, 16px and 96px). |
| `build/icon.svg`, `build/icon.png` | Icon source and the electron-builder application icon (512px). |
| `docs/configuration.md` | User-facing reference for every config option. |
| `tests/unit/*.test.js` | Unit tests, one file per unit under test. |

---

### Task 1: Project scaffolding

Creates the repo skeleton so every later task has a working `npm run lint` and `npm run test:unit`.

**Files:**
- Create: `package.json`
- Create: `eslint.config.mjs`
- Create: `.gitignore`
- Create: `LICENSE`
- Create: `README.md`
- Test: `tests/unit/scaffolding.test.js`

**Interfaces:**
- Consumes: nothing (first task).
- Produces: `npm run lint`, `npm run test:unit`, `npm start`, `npm run pack`, `npm run dist:linux` scripts. Dependency set: runtime `electron-log`, `yargs`; dev `electron`, `electron-builder`, `eslint`, `@eslint/js`, `globals`. The `main` entry point is `app/index.js`.

> **Deliberate omissions vs teams-for-linux (YAGNI):** no `electron-store` (the MVP has no persisted settings — config is read-only at startup), no `electron-window-state` (window geometry persistence is not in the spec's success criteria), no `playwright` (E2E is explicitly deferred). Do not add them.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/scaffolding.test.js`:

```javascript
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/scaffolding.test.js`
Expected: FAIL — `Cannot find module '.../package.json'`.

- [ ] **Step 3: Create `package.json`**

```json
{
  "name": "outlook-for-linux",
  "version": "0.1.0",
  "main": "app/index.js",
  "description": "Unofficial client for Microsoft Outlook for Linux",
  "keywords": [
    "Outlook",
    "Microsoft Outlook"
  ],
  "license": "GPL-3.0-or-later",
  "scripts": {
    "lint": "eslint app tests",
    "test:unit": "node --test 'tests/unit/*.test.js'",
    "start": "electron . --trace-warnings",
    "start:dev": "electron . --trace-warnings --no-sandbox",
    "pack": "electron-builder --dir",
    "dist": "electron-builder",
    "dist:linux": "electron-builder --linux"
  },
  "dependencies": {
    "electron-log": "^5.4.4",
    "yargs": "^18.1.0"
  },
  "devDependencies": {
    "@eslint/js": "^10.0.1",
    "electron": "42.8.1",
    "electron-builder": "26.15.7",
    "eslint": "^10.8.0",
    "globals": "^17.9.0"
  },
  "build": {
    "appId": "outlook-for-linux",
    "linux": {
      "category": "Office;Network;Email",
      "packageCategory": "net",
      "executableName": "outlook-for-linux",
      "synopsis": "Outlook for Linux",
      "description": "Unofficial Microsoft Outlook client for Linux using Electron. It wraps the Outlook web app as a standalone application.",
      "icon": "build/icon.png",
      "desktop": {
        "entry": {
          "Name": "Outlook for Linux",
          "Comment": "Unofficial client for Microsoft Outlook for Linux"
        }
      },
      "target": [
        "AppImage",
        "deb",
        "rpm"
      ]
    },
    "rpm": {
      "depends": [
        "gtk3",
        "libnotify",
        "nss",
        "libXScrnSaver",
        "(libXtst or libXtst6)",
        "xdg-utils",
        "at-spi2-core",
        "(libuuid or libuuid1)"
      ]
    }
  }
}
```

- [ ] **Step 4: Create `eslint.config.mjs`**

```javascript
import globals from "globals";
import pluginJs from "@eslint/js";

export default [
  { files: ["**/*.js"], languageOptions: { sourceType: "commonjs" } },
  { languageOptions: { globals: { ...globals.browser, ...globals.node } } },
  pluginJs.configs.recommended,
  {
    rules: {
      "no-var": "error",
      eqeqeq: ["error", "always"],
    },
  },
];
```

- [ ] **Step 5: Create `.gitignore`**

```gitignore
node_modules/
dist/
out/
*.log
.DS_Store
```

- [ ] **Step 6: Add the GPL-3.0 licence text**

Run — this copies the canonical text already present in the sibling checkout rather than hand-typing it:

```bash
cp /home/sandesh/myfolder/projects/teams-for-linux/LICENSE.md LICENSE
```

If that path does not exist, fetch it instead: `curl -fsSL https://www.gnu.org/licenses/gpl-3.0.txt -o LICENSE`.

- [ ] **Step 7: Create `README.md`**

```markdown
# Outlook for Linux

Unofficial desktop client for the Microsoft Outlook web app, built with Electron.
Modelled on [teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux).

## Status

MVP under development. See `docs/superpowers/specs/` for the design and
`docs/superpowers/plans/` for the implementation plan.

## Development

```bash
npm install
npm start          # run the app
npm run lint       # ESLint
npm run test:unit  # unit tests
```

## Configuration

Create `~/.config/outlook-for-linux/config.json`:

```json
{
  "url": "https://outlook.office.com",
  "trayIconEnabled": true
}
```

Any option can also be passed as a CLI flag, e.g. `--url=https://outlook.live.com`.
CLI flags take precedence over the config file, which takes precedence over defaults.

## Licence

GPL-3.0-or-later.
```

- [ ] **Step 8: Install dependencies**

Run: `npm install`
Expected: completes and creates `package-lock.json` and `node_modules/`.

- [ ] **Step 9: Run the test to verify it passes**

Run: `node --test tests/unit/scaffolding.test.js`
Expected: PASS — 5 tests.

- [ ] **Step 10: Verify lint runs**

Run: `npm run lint`
Expected: exit 0. (`app/` is empty at this point, which is fine — ESLint reports no errors.)

- [ ] **Step 11: Commit**

```bash
git add package.json package-lock.json eslint.config.mjs .gitignore LICENSE README.md tests/unit/scaffolding.test.js
git commit -m "feat: scaffold outlook-for-linux project"
```

---

### Task 2: PII log sanitizer

The lowest-level utility: every other module logs through it, so it lands first.

**Files:**
- Create: `app/utils/logSanitizer.js`
- Test: `tests/unit/logSanitizer.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `module.exports = { sanitize, sanitizeObject, sanitizeLogData, containsPII, PII_PATTERNS }` where
  - `sanitize(message: any) => string` — redacts PII from a single value, stringifying non-strings.
  - `sanitizeObject(obj: any, seen?: WeakSet) => any` — recursively redacts strings inside objects/arrays/Errors, replaces values of keys containing `password`/`token`/`secret`/`key` with `"[REDACTED]"`, and returns `"[Circular]"` for cycles.
  - `sanitizeLogData(messageData: Array) => Array` — maps an `electron-log` argument array through the above. Used by `app/config/logger.js`.
  - `containsPII(message: any) => boolean` — true if any pattern matches.
  - `PII_PATTERNS` — the raw pattern map, exported for tests.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/logSanitizer.test.js`:

```javascript
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/logSanitizer.test.js`
Expected: FAIL — `Cannot find module '../../app/utils/logSanitizer'`.

- [ ] **Step 3: Port the implementation**

Copy the module from the sibling checkout — it is generic (no Teams-specific content) and battle-tested, so port it unchanged rather than rewriting:

```bash
mkdir -p app/utils
cp /home/sandesh/myfolder/projects/teams-for-linux/app/utils/logSanitizer.js app/utils/logSanitizer.js
```

Then make one deliberate trim: delete the `mqttUrl` pattern from `PII_PATTERNS` and its corresponding `replaceAll` line in `sanitize()`, since MQTT is out of scope for the MVP. Leave every other pattern in place — `email`, `uuid`, `password`, `bearerToken`, `ipAddress`, `urlQueryParams`, `authHeader`, `apiKey`, `accessToken`, `refreshToken`, `clientSecret`, `certFingerprint`, `userPath`.

Also delete the `createSanitizer` and `detectPIITypes` exports and their function bodies — nothing in this MVP calls them (YAGNI). Keep the module's `'use strict';` header, tab indentation, and the ordering comment (`// Order matters: specific patterns before general ones`).

If the sibling checkout is unavailable, the module must implement exactly the exported interface listed above, with `sanitize` applying replacements in this order: bearer token, password, auth header, api key, access token, refresh token, client secret, cert fingerprint, email, UUID, IP address, URL query params, user path.

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/logSanitizer.test.js`
Expected: PASS — 11 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/utils/logSanitizer.js tests/unit/logSanitizer.test.js
git commit -m "feat: add PII log sanitizer"
```

---

### Task 3: Config defaults and option schema

Two pure data modules — no Electron imports — so tests and any future docs generator can load them outside Electron.

**Files:**
- Create: `app/config/defaults.js`
- Create: `app/config/options.js`
- Test: `tests/unit/configOptions.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `app/config/defaults.js` → `module.exports = { NETWORK_ERROR_PATTERNS }` where `NETWORK_ERROR_PATTERNS: string[]`. Consumed by `app/utils/networkErrors.js` (Task 4).
  - `app/config/options.js` → `module.exports` is a yargs options object. Every entry has `{ default, describe, type }`; `appIconType` also has `choices`. Consumed by `app/config/index.js` (Task 6). Option names, types and defaults: `url` (string, `"https://outlook.office.com"`), `appTitle` (string, `"Microsoft Outlook"`), `appIcon` (string, `""`), `appIconType` (string, `"default"`, choices `["default", "light", "dark"]`), `trayIconEnabled` (boolean, `true`), `closeToTray` (boolean, `true`), `disableNotifications` (boolean, `false`), `defaultNotificationUrgency` (string, `"normal"`, choices `["low", "normal", "critical"]`), `useMutationTitleLogic` (boolean, `true`), `disableBadgeCount` (boolean, `false`), `disableNotificationWindowFlash` (boolean, `false`), `disableGpu` (boolean, `false`), `logConfig` (object, `{ transports: { console: { level: "info" }, file: { level: false } } }`), `watchConfigFile` (boolean, `false`), `webDebug` (boolean, `false`).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/configOptions.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const options = require("../../app/config/options");
const { NETWORK_ERROR_PATTERNS } = require("../../app/config/defaults");

test("every option declares a default, a description and a type", () => {
  for (const [name, option] of Object.entries(options)) {
    assert.ok("default" in option, `${name} has no default`);
    assert.ok(option.describe, `${name} has no describe`);
    assert.ok(option.type, `${name} has no type`);
  }
});

test("the Outlook URL defaults to the work/school endpoint", () => {
  assert.strictEqual(options.url.default, "https://outlook.office.com");
  assert.strictEqual(options.url.type, "string");
});

test("tray and notifications are on by default", () => {
  assert.strictEqual(options.trayIconEnabled.default, true);
  assert.strictEqual(options.disableNotifications.default, false);
  assert.strictEqual(options.useMutationTitleLogic.default, true);
});

test("appIconType is constrained to the shipped icon variants", () => {
  assert.deepStrictEqual(options.appIconType.choices, [
    "default",
    "light",
    "dark",
  ]);
  assert.ok(options.appIconType.choices.includes(options.appIconType.default));
});

test("options.js is loadable without Electron", () => {
  // A require cycle through Electron would have thrown above; assert the
  // module stayed pure data.
  assert.strictEqual(typeof options, "object");
  assert.ok(!Array.isArray(options));
});

test("no Teams-specific or out-of-scope options leaked in", () => {
  for (const name of Object.keys(options)) {
    assert.ok(!/teams|mqtt|graph|meetup|spellcheck/i.test(name), `unexpected option: ${name}`);
  }
});

test("network error patterns cover the transient connection failures", () => {
  for (const pattern of [
    "ERR_INTERNET_DISCONNECTED",
    "ERR_NAME_NOT_RESOLVED",
    "ERR_CONNECTION_RESET",
    "ERR_NETWORK_CHANGED",
  ]) {
    assert.ok(NETWORK_ERROR_PATTERNS.includes(pattern), `missing ${pattern}`);
  }
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/configOptions.test.js`
Expected: FAIL — `Cannot find module '../../app/config/options'`.

- [ ] **Step 3: Create `app/config/defaults.js`**

```javascript
/**
 * Default values shared across modules. Kept free of Electron imports so unit
 * tests (and any future docs generator) can require it outside Electron.
 */

// Network error patterns that indicate transient connection issues
// (proxy, tunnel, DNS, etc.). Used by app/utils/networkErrors.js so the global
// error handlers do not terminate the app on a network blip.
const NETWORK_ERROR_PATTERNS = [
  "ERR_TUNNEL_CONNECTION_FAILED",
  "ERR_PROXY_CONNECTION_FAILED",
  "ERR_INTERNET_DISCONNECTED",
  "ERR_NETWORK_CHANGED",
  "ERR_CONNECTION_RESET",
  "ERR_CONNECTION_REFUSED",
  "ERR_CONNECTION_TIMED_OUT",
  "ERR_NAME_NOT_RESOLVED",
];

module.exports = { NETWORK_ERROR_PATTERNS };
```

- [ ] **Step 4: Create `app/config/options.js`**

```javascript
// Configuration option definitions for Outlook for Linux.
//
// Single source of truth for the wrapper's config surface. Consumed at runtime
// by the yargs parser in ./index.js. Keep it a plain data module with no
// imports so it stays loadable outside Electron.

module.exports = {
  url: {
    default: "https://outlook.office.com",
    describe:
      "Outlook web URL. Use https://outlook.live.com for personal accounts, or a tenant-specific URL.",
    type: "string",
  },
  appTitle: {
    default: "Microsoft Outlook",
    describe: "Window and tray tooltip title",
    type: "string",
  },
  appIcon: {
    default: "",
    describe: "Absolute path to a custom tray icon. Empty means use the bundled icon.",
    type: "string",
  },
  appIconType: {
    default: "default",
    describe: "Which bundled tray icon variant to use",
    type: "string",
    choices: ["default", "light", "dark"],
  },
  trayIconEnabled: {
    default: true,
    describe: "Enable the tray icon",
    type: "boolean",
  },
  closeToTray: {
    default: true,
    describe: "Hide to tray instead of quitting when the window is closed",
    type: "boolean",
  },
  disableNotifications: {
    default: false,
    describe: "Disable all native notifications",
    type: "boolean",
  },
  defaultNotificationUrgency: {
    default: "normal",
    describe: "Urgency passed to the native notification server",
    type: "string",
    choices: ["low", "normal", "critical"],
  },
  useMutationTitleLogic: {
    default: true,
    describe:
      "Watch the page title with a MutationObserver to derive the unread count for the tray badge",
    type: "boolean",
  },
  disableBadgeCount: {
    default: false,
    describe: "Do not draw the unread count badge on the tray icon",
    type: "boolean",
  },
  disableNotificationWindowFlash: {
    default: false,
    describe: "Do not flash the window frame when the unread count increases",
    type: "boolean",
  },
  disableGpu: {
    default: false,
    describe: "Disable GPU hardware acceleration",
    type: "boolean",
  },
  logConfig: {
    default: {
      transports: {
        console: { level: "info" },
        file: { level: false },
      },
    },
    describe:
      'electron-log configuration. Set to the string "console" for plain console logging, or false to disable logging.',
    type: "object",
  },
  watchConfigFile: {
    default: false,
    describe: "Reload the app when the config file changes",
    type: "boolean",
  },
  webDebug: {
    default: false,
    describe: "Open DevTools on start",
    type: "boolean",
  },
};
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `node --test tests/unit/configOptions.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 6: Lint and commit**

```bash
npm run lint
git add app/config/defaults.js app/config/options.js tests/unit/configOptions.test.js
git commit -m "feat: add config defaults and option schema"
```

---

### Task 4: Network error classifier

Isolated so the global error handlers in `app/index.js` (Task 17) stay one-liners and this decision is unit-tested.

**Files:**
- Create: `app/utils/networkErrors.js`
- Test: `tests/unit/networkErrors.test.js`

**Interfaces:**
- Consumes: `NETWORK_ERROR_PATTERNS` from `app/config/defaults.js` (Task 3).
- Produces: `module.exports = { isNetworkError }` where `isNetworkError(message: any) => boolean`. Returns `false` for non-strings. Consumed by `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/networkErrors.test.js`:

```javascript
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/networkErrors.test.js`
Expected: FAIL — `Cannot find module '../../app/utils/networkErrors'`.

- [ ] **Step 3: Create `app/utils/networkErrors.js`**

```javascript
const { NETWORK_ERROR_PATTERNS } = require("../config/defaults");

/**
 * Classifies an error message as a transient network failure.
 *
 * The global error handlers use this to avoid terminating the app on a
 * network blip: a dropped VPN or a DNS hiccup surfaces as an uncaught
 * exception from Chromium, which is not a reason to kill the process.
 *
 * @param {any} message - The error message to classify.
 * @returns {boolean} True when the message indicates a transient failure.
 */
function isNetworkError(message) {
  if (typeof message !== "string") return false;
  if (NETWORK_ERROR_PATTERNS.some((pattern) => message.includes(pattern))) {
    return true;
  }
  // "Object has been destroyed" occurs when the window is torn down during a
  // network-triggered operation (e.g. a reload after network recovery).
  if (message.includes("Object has been destroyed")) return true;
  // "Script failed to execute" occurs when executeJavaScript runs on a page
  // where the APIs are unavailable, e.g. a Chrome error page after
  // ERR_NAME_NOT_RESOLVED. A symptom of network failure, not a fatal error.
  if (message.includes("Script failed to execute")) return true;
  return false;
}

module.exports = { isNetworkError };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/networkErrors.test.js`
Expected: PASS — 5 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/utils/networkErrors.js tests/unit/networkErrors.test.js
git commit -m "feat: add transient network error classifier"
```

---

### Task 5: Logger initialisation with the PII hook

Wires `electron-log` and installs the sanitizer hook so no transport can emit PII.

**Files:**
- Create: `app/config/logger.js`
- Test: `tests/unit/logger.test.js`

**Interfaces:**
- Consumes: `sanitizeLogData` from `app/utils/logSanitizer.js` (Task 2).
- Produces: `module.exports = { init }` where `init(config: object|string|false, log?: object) => void`.
  - `config === "console"` — leave the global `console` alone.
  - `config` falsy — replace `console.log/info/debug/warn/error` with no-ops.
  - otherwise — merge `config` into `log`, call `log.initialize()`, push the sanitizing hook onto `log.hooks`, and `Object.assign(console, log.functions)`.
  - The second parameter exists for testability and defaults to `require("electron-log/main")`; production callers pass one argument. Consumed by `app/config/index.js` (Task 6).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/logger.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const logger = require("../../app/config/logger");

function fakeLog() {
  return {
    hooks: [],
    functions: { log() {}, info() {}, debug() {}, warn() {}, error() {} },
    transports: { console: { level: "info" }, file: { level: false } },
    initialized: false,
    initialize() {
      this.initialized = true;
    },
  };
}

test("installs a hook that strips PII from log data", () => {
  const log = fakeLog();
  logger.init({ transports: { console: { level: "debug" } } }, log);

  assert.strictEqual(log.initialized, true);
  assert.strictEqual(log.hooks.length, 1);

  const message = { data: ["contact frank@contoso.com", { token: "abc" }] };
  const hooked = log.hooks[0](message);
  assert.deepStrictEqual(hooked.data, ["contact [EMAIL]", { token: "[REDACTED]" }]);
});

test("merges the supplied transport config into the log instance", () => {
  const log = fakeLog();
  logger.init({ transports: { console: { level: "debug" } } }, log);
  assert.strictEqual(log.transports.console.level, "debug");
  // Untouched keys survive the merge.
  assert.strictEqual(log.transports.file.level, false);
});

test('the "console" config leaves logging untouched', () => {
  const log = fakeLog();
  logger.init("console", log);
  assert.strictEqual(log.initialized, false);
  assert.strictEqual(log.hooks.length, 0);
});

test("a falsy config silences the console", () => {
  const original = { ...console };
  try {
    logger.init(false, fakeLog());
    // No-op functions return undefined and must not throw.
    assert.strictEqual(console.info("swallowed"), undefined);
    assert.strictEqual(console.error("swallowed"), undefined);
  } finally {
    Object.assign(console, original);
  }
});

test("merging refuses prototype-polluting keys", () => {
  const log = fakeLog();
  const malicious = JSON.parse('{"__proto__": {"polluted": true}}');
  logger.init(malicious, log);
  assert.strictEqual({}.polluted, undefined);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/logger.test.js`
Expected: FAIL — `Cannot find module '../../app/config/logger'`.

- [ ] **Step 3: Create `app/config/logger.js`**

```javascript
const { sanitizeLogData } = require("../utils/logSanitizer");

const FORBIDDEN_KEYS = new Set(["__proto__", "constructor", "prototype"]);

/**
 * Deep-merges source into target, preserving functions on the target by
 * assigning source properties onto them (electron-log's transports are
 * callable objects carrying configuration properties).
 */
function mergeWith(target, source) {
  for (const key of Object.keys(source)) {
    if (FORBIDDEN_KEYS.has(key)) continue;

    const targetValue = target[key];
    const sourceValue = source[key];

    if (typeof targetValue === "function") {
      Object.assign(targetValue, sourceValue);
    } else if (
      sourceValue &&
      typeof sourceValue === "object" &&
      !Array.isArray(sourceValue) &&
      targetValue &&
      typeof targetValue === "object" &&
      !Array.isArray(targetValue)
    ) {
      mergeWith(targetValue, sourceValue);
    } else {
      target[key] = sourceValue;
    }
  }
  return target;
}

function silenceConsole() {
  console.log = function () {};
  console.info = function () {};
  console.debug = function () {};
  console.warn = function () {};
  console.error = function () {};
}

/**
 * Initialises logging.
 *
 * @param {object|string|false} config - electron-log config, the literal
 *   string "console" to keep the plain console, or a falsy value to silence
 *   logging entirely.
 * @param {object} [log] - The electron-log instance. Injectable for tests.
 */
exports.init = function (config, log = require("electron-log/main")) {
  if (!config) {
    console.info("[LOGGER] Disabling logs");
    silenceConsole();
    return;
  }

  if (config === "console") {
    console.debug("[LOGGER] Using the default console");
    return;
  }

  mergeWith(log, config);
  log.initialize();

  // PII sanitisation applies to every transport, so nothing can leak through
  // a transport added later.
  log.hooks.push((message) => {
    message.data = sanitizeLogData(message.data);
    return message;
  });

  Object.assign(console, log.functions);
  console.debug("[LOGGER] Logger initialised");
};
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/logger.test.js`
Expected: PASS — 5 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/config/logger.js tests/unit/logger.test.js
git commit -m "feat: initialise electron-log with a PII sanitizing hook"
```

---

### Task 6: Config loader with defaults → file → CLI precedence

The heart of the configuration system, and the spec's central success criterion for config. Deliberately free of Electron imports so precedence is tested for real: the caller supplies the config directory and the argv array.

**Files:**
- Create: `app/config/index.js`
- Test: `tests/unit/config.test.js`

**Interfaces:**
- Consumes: `app/config/options.js` (Task 3), `app/config/logger.js` (Task 5).
- Produces: `module.exports = argv` where `argv(configPath: string, appVersion: string, argvArray?: string[]) => object`.
  - Reads `<configPath>/config.json` if present; a malformed file is reported on the returned object's `error` property and otherwise ignored.
  - Precedence: option defaults, overridden by config-file values, overridden by `argvArray` (defaults to `process.argv.slice(1)`).
  - The returned object is deep-frozen. It always carries `configPath` (the directory it was loaded from) and `isConfigFile` (boolean).
  - Consumed by `app/appConfiguration/index.js` (Task 7).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/config.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const argv = require("../../app/config/index");

function withConfigDir(contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-config-"));
  if (contents !== undefined) {
    fs.writeFileSync(path.join(dir, "config.json"), contents);
  }
  return dir;
}

test("uses defaults when no config file exists", () => {
  const config = argv(withConfigDir(), "0.1.0", []);
  assert.strictEqual(config.url, "https://outlook.office.com");
  assert.strictEqual(config.trayIconEnabled, true);
  assert.strictEqual(config.isConfigFile, false);
});

test("the config file overrides defaults", () => {
  const dir = withConfigDir(
    JSON.stringify({ url: "https://outlook.live.com", trayIconEnabled: false }),
  );
  const config = argv(dir, "0.1.0", []);
  assert.strictEqual(config.url, "https://outlook.live.com");
  assert.strictEqual(config.trayIconEnabled, false);
  assert.strictEqual(config.isConfigFile, true);
});

test("CLI flags override the config file", () => {
  const dir = withConfigDir(JSON.stringify({ url: "https://outlook.live.com" }));
  const config = argv(dir, "0.1.0", ["--url=https://outlook.office365.com"]);
  assert.strictEqual(config.url, "https://outlook.office365.com");
});

test("CLI flags override defaults for booleans", () => {
  const config = argv(withConfigDir(), "0.1.0", ["--trayIconEnabled=false"]);
  assert.strictEqual(config.trayIconEnabled, false);
});

test("a malformed config file is reported and does not throw", () => {
  const dir = withConfigDir("{ this is not json");
  const config = argv(dir, "0.1.0", []);
  assert.ok(config.error, "expected an error to be reported");
  // Defaults still apply so the app can start.
  assert.strictEqual(config.url, "https://outlook.office.com");
});

test("the returned config is deeply frozen", () => {
  const config = argv(withConfigDir(), "0.1.0", []);
  assert.strictEqual(Object.isFrozen(config), true);
  assert.strictEqual(Object.isFrozen(config.logConfig), true);
  assert.throws(
    () => {
      "use strict";
      config.url = "https://evil.example.com";
    },
    TypeError,
  );
});

test("the config records where it was loaded from", () => {
  const dir = withConfigDir();
  const config = argv(dir, "0.1.0", []);
  assert.strictEqual(config.configPath, dir);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/config.test.js`
Expected: FAIL — `Cannot find module '../../app/config/index'`.

- [ ] **Step 3: Create `app/config/index.js`**

```javascript
const yargs = require("yargs");
const fs = require("node:fs");
const path = require("node:path");
const logger = require("./logger");
const configOptions = require("./options");

function getConfigFilePath(configPath) {
  return path.join(configPath, "config.json");
}

/**
 * Reads the user's config file, if any.
 *
 * @returns {{ configFile: object, isConfigFile: boolean, configError: string|null }}
 */
function readConfigFile(configPath) {
  const filePath = getConfigFilePath(configPath);

  if (!fs.existsSync(filePath)) {
    console.info("[CONFIG] No config file found, using defaults");
    return { configFile: {}, isConfigFile: false, configError: null };
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8"));
    console.info("[CONFIG] Loaded user configuration");
    return { configFile: parsed, isConfigFile: true, configError: null };
  } catch (e) {
    // The message can contain a file path, so it is sanitized by the logger
    // hook before it reaches any transport.
    console.warn(`[CONFIG] Error in config file, using defaults: ${e.message}`);
    return { configFile: {}, isConfigFile: false, configError: e.message };
  }
}

/** Recursively freezes an object so the config is immutable after startup. */
function deepFreeze(obj) {
  for (const value of Object.values(obj)) {
    if (value && typeof value === "object" && !Object.isFrozen(value)) {
      deepFreeze(value);
    }
  }
  return Object.freeze(obj);
}

/**
 * Resolves the application configuration.
 *
 * Precedence: option defaults < config file < CLI flags.
 *
 * @param {string} configPath - Directory holding config.json.
 * @param {string} appVersion - Version reported by --version.
 * @param {string[]} [argvArray] - Argument vector; defaults to the real one.
 * @returns {object} The deep-frozen configuration.
 */
function argv(configPath, appVersion, argvArray = process.argv.slice(1)) {
  const { configFile, isConfigFile, configError } = readConfigFile(configPath);

  // yargs v18 is no longer a singleton and requires explicit instantiation.
  const config = yargs()
    .config(configFile)
    .version(appVersion)
    .options(configOptions)
    .help()
    .parse(argvArray);

  if (configError) {
    config.error = configError;
  }
  config.configPath = configPath;
  config.isConfigFile = isConfigFile;

  logger.init(config.logConfig);

  console.info("[CONFIG] Configuration resolved");

  return deepFreeze(config);
}

module.exports = argv;
```

> **Note on the watcher:** teams-for-linux sets up `fs.watch` on the config file inside this module, which forces an `electron` import. Here the watcher lives in `app/index.js` (Task 17) instead, keyed off `config.watchConfigFile`, which keeps this module testable outside Electron.

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/config.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/config/index.js tests/unit/config.test.js
git commit -m "feat: resolve config from defaults, config file and CLI flags"
```

---

### Task 7: AppConfiguration wrapper

A thin class giving the rest of the app one object to depend on, with the resolved config behind a `#private` field so nothing can reassign it.

**Files:**
- Create: `app/appConfiguration/index.js`
- Test: `tests/unit/appConfiguration.test.js`

**Interfaces:**
- Consumes: `app/config/index.js` (Task 6).
- Produces: `module.exports = { AppConfiguration }`.
  - `new AppConfiguration(configPath: string, appVersion: string, argvArray?: string[])`
  - `get configPath(): string`
  - `get startupConfig(): object` — the frozen config from Task 6.
  - Consumed by `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/appConfiguration.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { AppConfiguration } = require("../../app/appConfiguration");

function tempConfigDir(contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-appconfig-"));
  if (contents) {
    fs.writeFileSync(path.join(dir, "config.json"), JSON.stringify(contents));
  }
  return dir;
}

test("exposes the resolved startup config", () => {
  const dir = tempConfigDir({ url: "https://outlook.live.com" });
  const appConfig = new AppConfiguration(dir, "0.1.0", []);
  assert.strictEqual(appConfig.startupConfig.url, "https://outlook.live.com");
  assert.strictEqual(appConfig.configPath, dir);
});

test("the startup config cannot be replaced from outside", () => {
  const appConfig = new AppConfiguration(tempConfigDir(), "0.1.0", []);
  assert.throws(() => {
    appConfig.startupConfig = { url: "https://evil.example.com" };
  }, TypeError);
});

test("private state is not enumerable on the instance", () => {
  const appConfig = new AppConfiguration(tempConfigDir(), "0.1.0", []);
  assert.deepStrictEqual(Object.keys(appConfig), []);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/appConfiguration.test.js`
Expected: FAIL — `Cannot find module '../../app/appConfiguration'`.

- [ ] **Step 3: Create `app/appConfiguration/index.js`**

```javascript
const argv = require("../config");

/**
 * Owns the application configuration for the process lifetime.
 *
 * The resolved config is read once at startup and kept behind a private field.
 * It is already deep-frozen by the config loader; treat it as immutable.
 */
class AppConfiguration {
  #configPath;
  #startupConfig;

  /**
   * @param {string} configPath - Directory holding config.json.
   * @param {string} appVersion - Version reported by --version.
   * @param {string[]} [argvArray] - Argument vector; defaults to the real one.
   */
  constructor(configPath, appVersion, argvArray) {
    this.#configPath = configPath;
    this.#startupConfig = argv(configPath, appVersion, argvArray);
  }

  get configPath() {
    return this.#configPath;
  }

  get startupConfig() {
    return this.#startupConfig;
  }
}

module.exports = { AppConfiguration };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/appConfiguration.test.js`
Expected: PASS — 3 tests. (The second test passes because a getter-only accessor throws on assignment in the strict-mode context of a CommonJS test file under `node --test`.)

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/appConfiguration/index.js tests/unit/appConfiguration.test.js
git commit -m "feat: add AppConfiguration wrapper"
```

---

### Task 8: IPC channel allowlist

The compensating control for a preload that shares the page's context. Only the channels this MVP actually uses are listed; the list grows with the app, never speculatively.

**Files:**
- Create: `app/security/ipcValidator.js`
- Test: `tests/unit/ipcValidator.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `module.exports = { validateIpcChannel, allowedChannels }`.
  - `allowedChannels: Set<string>` — exactly: `get-config`, `show-notification`, `notification-closed`, `tray-update`, `set-badge-count`, `unhandled-rejection`, `window-error`.
  - `validateIpcChannel(channel: string, payload?: any) => boolean` — false for channels off the list (with a warning), true otherwise; mutates `payload` in place to strip `__proto__` / `constructor` / `prototype` keys recursively.
  - Consumed by `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/ipcValidator.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const {
  validateIpcChannel,
  allowedChannels,
} = require("../../app/security/ipcValidator");

test("allows exactly the MVP channels", () => {
  assert.deepStrictEqual([...allowedChannels].sort(), [
    "get-config",
    "notification-closed",
    "set-badge-count",
    "show-notification",
    "tray-update",
    "unhandled-rejection",
    "window-error",
  ]);
});

test("permits an allowlisted channel", () => {
  assert.strictEqual(validateIpcChannel("show-notification", { title: "hi" }), true);
});

test("blocks an unknown channel", () => {
  assert.strictEqual(validateIpcChannel("rm-rf-slash"), false);
});

test("blocks Teams-era channels that were not ported", () => {
  for (const channel of ["choose-desktop-media", "get-teams-settings", "graph-api-get-mail-messages"]) {
    assert.strictEqual(validateIpcChannel(channel), false, `${channel} should be blocked`);
  }
});

test("strips prototype-pollution keys from the payload", () => {
  const payload = JSON.parse('{"title":"hi","__proto__":{"polluted":true}}');
  assert.strictEqual(validateIpcChannel("show-notification", payload), true);
  assert.strictEqual({}.polluted, undefined);
  assert.strictEqual(Object.hasOwn(payload, "__proto__"), false);
});

test("strips pollution keys from nested payloads", () => {
  const payload = JSON.parse('{"outer":{"inner":{"__proto__":{"bad":1}}}}');
  validateIpcChannel("tray-update", payload);
  assert.strictEqual(Object.hasOwn(payload.outer.inner, "__proto__"), false);
});

test("tolerates a missing payload and cyclic payloads", () => {
  assert.strictEqual(validateIpcChannel("get-config"), true);
  const cyclic = { a: 1 };
  cyclic.self = cyclic;
  assert.strictEqual(validateIpcChannel("get-config", cyclic), true);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/ipcValidator.test.js`
Expected: FAIL — `Cannot find module '../../app/security/ipcValidator'`.

- [ ] **Step 3: Create `app/security/ipcValidator.js`**

```javascript
/**
 * IPC security validation.
 *
 * The preload script shares the Outlook page's context, so the main process
 * treats every inbound channel name as untrusted: it must appear on this
 * allowlist, and its payload is stripped of prototype-pollution vectors.
 *
 * Add a channel here in the same change that registers its handler.
 */
const allowedChannels = new Set([
  // Renderer asks main for the resolved startup config (invoke).
  "get-config",

  // Notifications: renderer forwards Outlook's web notification (invoke),
  // main tells the renderer the OS dismissed it (main → renderer; not gated
  // by this validator, listed so the allowlist stays authoritative).
  "show-notification",
  "notification-closed",

  // Tray: renderer pushes a rendered badge icon (send) and the unread count
  // for the OS badge (invoke).
  "tray-update",
  "set-badge-count",

  // Renderer-side error forwarding, registered in app/browser/preload.js.
  "unhandled-rejection",
  "window-error",
]);

const DANGEROUS_PROPS = new Set(["__proto__", "constructor", "prototype"]);
const MAX_SANITIZE_DEPTH = 10;

/**
 * Recursively removes prototype-pollution vectors from a payload, in place.
 *
 * @param {any} obj - The payload to sanitize.
 * @param {number} depth - Recursion depth; bounded so cycles cannot overflow.
 */
function sanitizePayload(obj, depth = 0) {
  if (!obj || typeof obj !== "object" || depth > MAX_SANITIZE_DEPTH) {
    return;
  }

  for (const prop of DANGEROUS_PROPS) {
    if (Object.hasOwn(obj, prop)) {
      delete obj[prop];
    }
  }

  for (const key of Object.keys(obj)) {
    if (obj[key] && typeof obj[key] === "object") {
      sanitizePayload(obj[key], depth + 1);
    }
  }
}

/**
 * Validates an inbound IPC request.
 *
 * @param {string} channel - The IPC channel name.
 * @param {any} [payload] - The payload, sanitized in place when present.
 * @returns {boolean} True when the request may proceed.
 */
function validateIpcChannel(channel, payload = null) {
  if (!allowedChannels.has(channel)) {
    console.warn(`[IPC] Blocked unauthorized channel: ${channel}`);
    return false;
  }

  sanitizePayload(payload);
  return true;
}

module.exports = { validateIpcChannel, allowedChannels };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/ipcValidator.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Run the whole suite and commit**

```bash
npm run test:unit
npm run lint
git add app/security/ipcValidator.js tests/unit/ipcValidator.test.js
git commit -m "feat: add IPC channel allowlist"
```

---

### Task 9: Command-line switches

**Files:**
- Create: `app/startup/commandLine.js`
- Test: `tests/unit/commandLine.test.js`

**Interfaces:**
- Consumes: the resolved config (Task 6).
- Produces: `module.exports = CommandLineManager` with one static method:
  - `CommandLineManager.applySwitches(config: object, electronApp: object) => void` — appends switches to `electronApp.commandLine`. Consumed by `app/index.js` (Task 17).

> **Simplification vs teams-for-linux:** teams-for-linux splits this into `addSwitchesBeforeConfigLoad` / `addSwitchesAfterConfigLoad` because its media and WebRTC switches must precede config load. This MVP has no media switches, so a single post-config call is enough. It also drops every WebRTC, screen-capture, autoplay, media-key, ANGLE, rasterization and `js-flags` switch — none of them matter to a mail client.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/commandLine.test.js`:

```javascript
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/commandLine.test.js`
Expected: FAIL — `Cannot find module '../../app/startup/commandLine'`.

- [ ] **Step 3: Create `app/startup/commandLine.js`**

```javascript
const APP_CLASS = "outlook-for-linux";

/**
 * Applies Chromium/Electron command-line switches derived from config.
 *
 * Must run before the first window is created.
 */
class CommandLineManager {
  /**
   * @param {object} config - The resolved startup config.
   * @param {object} electronApp - Electron's `app` object. Injectable for tests.
   */
  static applySwitches(config, electronApp) {
    // Both switches are needed for WM_CLASS to reach X11 and Wayland
    // compositors, which use it to match the window to its .desktop entry
    // (correct icon and taskbar grouping).
    electronApp.commandLine.appendSwitch("class", APP_CLASS);
    electronApp.commandLine.appendSwitch("wm-class", APP_CLASS);

    if (config.disableGpu) {
      console.info("[STARTUP] Disabling GPU acceleration");
      electronApp.commandLine.appendSwitch("disable-gpu");
      electronApp.commandLine.appendSwitch("disable-gpu-compositing");
      electronApp.commandLine.appendSwitch("disable-software-rasterizer");
    }
  }
}

module.exports = CommandLineManager;
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/commandLine.test.js`
Expected: PASS — 4 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/startup/commandLine.js tests/unit/commandLine.test.js
git commit -m "feat: apply command-line switches from config"
```

---

### Task 10: Tray icon assets and chooser

Generates the icon set and the module that picks one. Both the tray (Task 13) and the badge renderer (Task 12) depend on these files existing.

**Files:**
- Create: `build/icon.svg`
- Create: `build/icon-monochrome.svg` (source for the panel variants — see the deviation note in Step 2)
- Create: `build/icon.png` (512×512, generated)
- Create: `app/assets/icons/icon-16x16.png`, `icon-96x96.png`, `icon-monochrome-light-16x16.png`, `icon-monochrome-light-96x96.png`, `icon-monochrome-dark-16x16.png`, `icon-monochrome-dark-96x96.png` (generated)
- Create: `app/browser/tools/trayIconChooser.js`
- Test: `tests/unit/trayIconChooser.test.js`

**Interfaces:**
- Consumes: `config.appIcon` and `config.appIconType` (Task 3).
- Produces: `module.exports = TrayIconChooser` (a class, not an instance).
  - `new TrayIconChooser(config: object)`
  - `getFile(): string` — an absolute path. Returns `config.appIcon` when it is a non-empty string after trimming; otherwise the bundled icon for `config.appIconType` at 96px.
  - Consumed by `app/browser/tools/trayIconRenderer.js` (Task 12) and `app/index.js` (Task 17).

> **Trademark note:** do not copy Microsoft's Outlook logo or teams-for-linux's Teams icons. This is an unofficial client, so the icon is a generic envelope glyph authored here.

- [ ] **Step 1: Create `build/icon.svg`**

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96" width="96" height="96">
  <rect x="6" y="6" width="84" height="84" rx="16" fill="#0f6cbd"/>
  <path d="M22 34h52v30a4 4 0 0 1-4 4H26a4 4 0 0 1-4-4z" fill="#ffffff"/>
  <path d="M22 34l26 20 26-20" fill="none" stroke="#0f6cbd" stroke-width="5"
        stroke-linejoin="round" stroke-linecap="round"/>
</svg>
```

- [ ] **Step 2: Generate the raster icons**

ImageMagick 7 (`magick`) is available on this machine and rasterizes this SVG directly.

```bash
mkdir -p app/assets/icons build

# Full-colour application and tray icons
magick -background none build/icon.svg -resize 512x512 build/icon.png
magick -background none build/icon.svg -resize 96x96 app/assets/icons/icon-96x96.png
magick -background none build/icon.svg -resize 16x16 app/assets/icons/icon-16x16.png

# Monochrome variants. "light" is the white glyph for dark panels; "dark" is
# the black glyph for light panels. Alpha is preserved so the panel shows
# through.
for size in 16 96; do
  magick -background none build/icon-monochrome.svg -resize ${size}x${size} \
    app/assets/icons/icon-monochrome-light-${size}x${size}.png
  magick app/assets/icons/icon-monochrome-light-${size}x${size}.png \
    -channel RGB -negate \
    app/assets/icons/icon-monochrome-dark-${size}x${size}.png
done
```

> **Deviation, applied during implementation:** the monochrome variants are
> generated from a second source file, `build/icon-monochrome.svg` (a stroked
> envelope glyph on a transparent background), not from `build/icon.svg`.
> Flattening the full-colour icon with `-alpha extract -background white
> -alpha shape` collapses it to its silhouette — a rounded square with
> `unique_gray=1` and no glyph — which reads as a blank block in the panel.
> Verify polarity rather than just file size:
>
> ```bash
> magick app/assets/icons/icon-monochrome-light-96x96.png \
>   -background black -alpha remove -format "%[fx:mean]\n" info:  # ~0.31
> magick app/assets/icons/icon-monochrome-dark-96x96.png \
>   -background white -alpha remove -format "%[fx:mean]\n" info:  # ~0.69
> ```

- [ ] **Step 3: Verify the generated files**

Run: `magick identify build/icon.png app/assets/icons/*.png`
Expected: seven PNGs, with the dimensions encoded in each filename (and `512x512` for `build/icon.png`). If any file is 0 bytes or missing, the SVG delegate failed — regenerate that file with the PIL fallback:

```bash
python3 -c "
import sys
from PIL import Image
src = Image.open('app/assets/icons/icon-96x96.png').convert('RGBA')
for size in (16, 96):
    src.resize((size, size), Image.LANCZOS).save(f'app/assets/icons/icon-{size}x{size}.png')
"
```

- [ ] **Step 4: Write the failing test**

Create `tests/unit/trayIconChooser.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const TrayIconChooser = require("../../app/browser/tools/trayIconChooser");

test("returns the bundled default icon when no custom icon is set", () => {
  const file = new TrayIconChooser({ appIcon: "", appIconType: "default" }).getFile();
  assert.ok(path.isAbsolute(file));
  assert.strictEqual(path.basename(file), "icon-96x96.png");
  assert.ok(fs.existsSync(file), "the bundled icon must exist on disk");
});

test("selects the monochrome variants by icon type", () => {
  for (const [type, expected] of [
    ["light", "icon-monochrome-light-96x96.png"],
    ["dark", "icon-monochrome-dark-96x96.png"],
  ]) {
    const file = new TrayIconChooser({ appIcon: "", appIconType: type }).getFile();
    assert.strictEqual(path.basename(file), expected);
    assert.ok(fs.existsSync(file), `${expected} must exist on disk`);
  }
});

test("a custom icon path wins over the bundled icons", () => {
  const custom = "/opt/icons/my-outlook.png";
  const file = new TrayIconChooser({ appIcon: custom, appIconType: "default" }).getFile();
  assert.strictEqual(file, custom);
});

test("a whitespace-only custom icon path falls back to the bundled icon", () => {
  const file = new TrayIconChooser({ appIcon: "   ", appIconType: "default" }).getFile();
  assert.strictEqual(path.basename(file), "icon-96x96.png");
});

test("an unknown icon type falls back to the default icon", () => {
  const file = new TrayIconChooser({ appIcon: "", appIconType: "chartreuse" }).getFile();
  assert.strictEqual(path.basename(file), "icon-96x96.png");
});
```

- [ ] **Step 5: Run the test to verify it fails**

Run: `node --test tests/unit/trayIconChooser.test.js`
Expected: FAIL — `Cannot find module '../../app/browser/tools/trayIconChooser'`.

- [ ] **Step 6: Create `app/browser/tools/trayIconChooser.js`**

```javascript
const path = require("node:path");

const iconFolder = path.join(__dirname, "../..", "assets/icons");

const icons = {
  default: "icon-96x96.png",
  light: "icon-monochrome-light-96x96.png",
  dark: "icon-monochrome-dark-96x96.png",
};

/** Resolves which tray icon file to use, from config. */
class TrayIconChooser {
  #config;

  /** @param {object} config - The resolved startup config. */
  constructor(config) {
    this.#config = config;
  }

  /** @returns {string} An absolute path to the icon file. */
  getFile() {
    const custom = this.#config.appIcon;
    if (typeof custom === "string" && custom.trim() !== "") {
      return custom;
    }
    const file = icons[this.#config.appIconType] ?? icons.default;
    return path.join(iconFolder, file);
  }
}

module.exports = TrayIconChooser;
```

- [ ] **Step 7: Run the test to verify it passes**

Run: `node --test tests/unit/trayIconChooser.test.js`
Expected: PASS — 5 tests.

- [ ] **Step 8: Lint and commit**

```bash
npm run lint
git add build app/assets/icons app/browser/tools/trayIconChooser.js tests/unit/trayIconChooser.test.js
git commit -m "feat: add tray icon assets and chooser"
```

---

### Task 11: Unread count from the page title

The most DOM-fragile piece in the app. It watches `<title>` for Outlook's leading `(N)` unread marker and republishes it as an `unread-count` DOM event, which the badge renderer (Task 12) consumes.

**Files:**
- Create: `app/browser/tools/mutationTitle.js`
- Test: `tests/unit/mutationTitle.test.js`

**Interfaces:**
- Consumes: `config.useMutationTitleLogic` (Task 3).
- Produces: `module.exports` is a singleton instance with:
  - `init(config: object) => void` — no-op unless `config.useMutationTitleLogic`; attaches a `MutationObserver` to `document.head` (waiting for `DOMContentLoaded` if the document is still loading).
  - `parseUnreadCount(title: any) => number|null` — exposed for testing and used internally. Returns the leading `(N)` count, `0` when there is no marker, and `null` when the input is unusable or out of the accepted `0…9999` range.
  - Dispatches `new CustomEvent("unread-count", { detail: { number } })` on `globalThis`, suppressing repeats of the same number.

> **PII:** the Outlook page title contains the signed-in user's display name and often their mailbox address. Never log the title — log only the derived number.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/mutationTitle.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const mutationTitle = require("../../app/browser/tools/mutationTitle");

test("extracts the leading unread marker", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("(3) Inbox - Grace Hopper - Outlook"),
    3,
  );
  assert.strictEqual(mutationTitle.parseUnreadCount("(147) Mail - Outlook"), 147);
});

test("reports zero when there is no marker", () => {
  assert.strictEqual(
    mutationTitle.parseUnreadCount("Inbox - Grace Hopper - Outlook"),
    0,
  );
});

test("ignores a parenthesised value that is not a leading count", () => {
  assert.strictEqual(mutationTitle.parseUnreadCount("Mail (draft) - Outlook"), 0);
  assert.strictEqual(mutationTitle.parseUnreadCount("Re: (3) budget - Outlook"), 0);
});

test("rejects unusable input", () => {
  assert.strictEqual(mutationTitle.parseUnreadCount(undefined), null);
  assert.strictEqual(mutationTitle.parseUnreadCount(null), null);
  assert.strictEqual(mutationTitle.parseUnreadCount(12345), null);
});

test("rejects implausible counts", () => {
  assert.strictEqual(mutationTitle.parseUnreadCount("(99999) Mail - Outlook"), null);
});

test("init is a no-op when the mutation title logic is disabled", () => {
  // No DOM exists in this test process, so a no-op is observable as
  // "did not throw".
  assert.doesNotThrow(() => mutationTitle.init({ useMutationTitleLogic: false }));
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/mutationTitle.test.js`
Expected: FAIL — `Cannot find module '../../app/browser/tools/mutationTitle'`.

- [ ] **Step 3: Create `app/browser/tools/mutationTitle.js`**

```javascript
const MAX_TITLE_LENGTH = 200;
const MAX_PLAUSIBLE_COUNT = 9999;
const UNREAD_MARKER = /^\((\d+)\)/;

/**
 * Derives the unread count from Outlook's page title.
 *
 * Outlook web prefixes the title with "(N)" when there are unread items, the
 * same convention Teams uses. This is best-effort: Outlook's DOM and title
 * format can change without notice, so every failure path degrades to
 * "no count" rather than throwing.
 */
class MutationObserverTitle {
  #lastNumber = -1;

  init(config) {
    if (!config.useMutationTitleLogic) {
      return;
    }
    console.debug("[TITLE] Unread-count observer enabled");

    if (globalThis.document?.readyState === "loading") {
      globalThis.addEventListener("DOMContentLoaded", () =>
        this.#attachObserver(),
      );
    } else {
      this.#attachObserver();
    }
  }

  /**
   * @param {any} title - The raw document title.
   * @returns {number|null} The unread count, or null when unusable.
   */
  parseUnreadCount(title) {
    if (typeof title !== "string") {
      return null;
    }

    const match = UNREAD_MARKER.exec(title.substring(0, MAX_TITLE_LENGTH));
    const number = match ? Number.parseInt(match[1], 10) : 0;

    if (Number.isNaN(number) || number < 0 || number > MAX_PLAUSIBLE_COUNT) {
      return null;
    }
    return number;
  }

  /** True when a mutation batch actually involves the <title> element. */
  #involvesTitleElement(mutations) {
    return mutations.some((mutation) => {
      const target = mutation.target;
      // characterData mutations target the text node inside <title>;
      // childList mutations on <title> target the element itself.
      if (
        target.nodeName === "TITLE" ||
        target.parentNode?.nodeName === "TITLE"
      ) {
        return true;
      }
      // <title> replaced outright, e.g. a React remount.
      const changedNodes = [...mutation.addedNodes, ...mutation.removedNodes];
      return changedNodes.some((node) => node.nodeName === "TITLE");
    });
  }

  #handleMutations(mutations) {
    try {
      if (!this.#involvesTitleElement(mutations)) {
        return;
      }

      // Never log the title itself: it carries the user's name and mailbox.
      const number = this.parseUnreadCount(globalThis.document.title);
      if (number === null || number === this.#lastNumber) {
        return;
      }
      this.#lastNumber = number;

      console.debug(`[TITLE] Unread count changed to ${number}`);
      globalThis.dispatchEvent(
        new CustomEvent("unread-count", { detail: { number } }),
      );
    } catch (error) {
      console.error("[TITLE] Observer callback failed:", error.message);
    }
  }

  #attachObserver() {
    try {
      if (!globalThis.document?.head || !globalThis.MutationObserver) {
        console.error("[TITLE] DOM environment unsuitable, unread count disabled");
        return;
      }

      const observer = new globalThis.MutationObserver((mutations) =>
        this.#handleMutations(mutations),
      );

      // Observing document.head with subtree survives Outlook replacing the
      // <title> element outright, and catches both childList (text node swap)
      // and characterData (in-place text edit) updates. Watching only the
      // current <title> with childList would miss in-place edits and leave the
      // badge stuck at a stale count.
      observer.observe(globalThis.document.head, {
        childList: true,
        characterData: true,
        subtree: true,
      });
      console.debug("[TITLE] Observer attached to document.head");
    } catch (error) {
      console.error("[TITLE] Failed to attach observer:", error.message);
    }
  }
}

module.exports = new MutationObserverTitle();
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/mutationTitle.test.js`
Expected: PASS — 6 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/browser/tools/mutationTitle.js tests/unit/mutationTitle.test.js
git commit -m "feat: derive unread count from the Outlook page title"
```

---

### Task 12: Tray badge renderer

Turns an `unread-count` event into a badge-composited tray icon and pushes it to main.

**Files:**
- Create: `app/browser/tools/trayIconRenderer.js`
- Test: `tests/unit/trayIconRenderer.test.js`

**Interfaces:**
- Consumes: `TrayIconChooser` (Task 10); `config.disableBadgeCount`, `config.disableNotificationWindowFlash` (Task 3).
- Produces: `module.exports` is a singleton instance with:
  - `init(config: object, ipcRenderer: object) => void` — loads the base icon via `nativeImage`, then listens for `unread-count` on `globalThis`.
  - `updateActivityCount(event: {detail: {number: number}}) => Promise<void>` — deduplicates, renders, and sends `tray-update` with `{ icon: string|null, flash: boolean, count: number }`, then invokes `set-badge-count` with the count unless `config.disableBadgeCount`.
  - `render(count: number) => Promise<string>` — a PNG data URL of the icon with the badge drawn on it. Overridden in tests.
  - Consumed by `app/browser/preload.js` (Task 15); its IPC is handled by `ApplicationTray` (Task 13) and `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/trayIconRenderer.test.js`. It drives the module without a DOM or a live Electron by setting the fields `init()` would have set and stubbing `render()`:

```javascript
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
    { channel: "tray-update", payload: { icon: FAKE_ICON, flash: true, count: 3 } },
  ]);
  assert.deepStrictEqual(invoked, [{ channel: "set-badge-count", payload: 3 }]);
});

test("clearing to zero sends a null icon so main restores the base icon", async () => {
  const { sent } = harness();
  await renderer.updateActivityCount({ detail: { number: 2 } });
  await renderer.updateActivityCount({ detail: { number: 0 } });

  assert.strictEqual(sent.length, 2);
  assert.deepStrictEqual(sent[1].payload, { icon: null, flash: false, count: 0 });
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/trayIconRenderer.test.js`
Expected: FAIL — `Cannot find module '../../app/browser/tools/trayIconRenderer'`.

- [ ] **Step 3: Create `app/browser/tools/trayIconRenderer.js`**

```javascript
const { nativeImage } = require("electron");
const TrayIconChooser = require("./trayIconChooser");

const CANVAS_SIZE = 140;
const BADGE_RADIUS = 40;
const BADGE_CENTRE_X = 100;
const BADGE_CENTRE_Y = 90;
const BADGE_TEXT_BASELINE_Y = 110;
const IMAGE_PNG = "image/png";

/**
 * Composites the unread count onto the tray icon and pushes it to the main
 * process. Runs in the renderer because it needs a canvas to draw the badge.
 */
class TrayIconRenderer {
  #lastRequestedCount;
  #updateSequence = 0;

  init(config, ipcRenderer) {
    this.ipcRenderer = ipcRenderer;
    this.config = config;

    const iconChooser = new TrayIconChooser(config);
    this.baseIcon = nativeImage.createFromPath(iconChooser.getFile());
    this.iconSize = this.baseIcon.getSize();

    globalThis.addEventListener(
      "unread-count",
      this.updateActivityCount.bind(this),
    );
    console.debug("[TRAY] Badge renderer initialised");
  }

  /** Clears deduplication state. Test seam only. */
  resetForTests() {
    this.#lastRequestedCount = undefined;
    this.#updateSequence = 0;
  }

  async updateActivityCount(event) {
    const count = event.detail.number;

    // Deduplicate against the most recently *requested* count, not the last
    // one that finished sending: comparing against the completed count would
    // swallow a clear-to-zero that arrives while a non-zero render is still
    // in flight, leaving the badge stuck.
    if (count === this.#lastRequestedCount) {
      return;
    }
    this.#lastRequestedCount = count;

    // Each update takes a sequence token, so an update whose render finishes
    // after a newer one started is discarded rather than overwriting it.
    const sequence = ++this.#updateSequence;

    // A zero count needs no canvas work: main falls back to the base icon.
    let icon = null;
    if (count > 0) {
      try {
        icon = await this.render(count);
      } catch (error) {
        console.error("[TRAY] Icon render failed:", error.message);
        // Allow a later event with the same count to retry.
        this.#lastRequestedCount = undefined;
        return;
      }
    }

    if (sequence !== this.#updateSequence) {
      return;
    }

    this.ipcRenderer.send("tray-update", {
      icon,
      flash: count > 0 && !this.config.disableNotificationWindowFlash,
      count,
    });

    if (!this.config.disableBadgeCount) {
      await this.ipcRenderer
        .invoke("set-badge-count", count)
        .catch((err) =>
          console.error("[TRAY] Failed to set badge count:", err.message),
        );
    }
  }

  /**
   * @param {number} count - The unread count to draw.
   * @returns {Promise<string>} A PNG data URL. Never rejects: on any drawing
   *   failure it resolves with the unbadged base icon.
   */
  render(count) {
    return new Promise((resolve) => {
      const baseIconData = this.baseIcon.toDataURL(IMAGE_PNG);
      if (!baseIconData || baseIconData === "data:,") {
        console.error("[TRAY] Base icon produced no data");
        resolve(baseIconData);
        return;
      }

      const canvas = document.createElement("canvas");
      canvas.width = CANVAS_SIZE;
      canvas.height = CANVAS_SIZE;

      const image = new Image();
      image.onerror = () => {
        console.error("[TRAY] Base icon failed to load for rendering");
        resolve(baseIconData);
      };
      image.onload = () => {
        try {
          resolve(this.#drawBadge(canvas, image, count));
        } catch (error) {
          console.error("[TRAY] Canvas drawing failed:", error.message);
          resolve(baseIconData);
        }
      };
      image.src = baseIconData;
    });
  }

  #drawBadge(canvas, image, count) {
    const ctx = canvas.getContext("2d");
    ctx.drawImage(image, 0, 0, CANVAS_SIZE, CANVAS_SIZE);

    if (count > 0 && !this.config.disableBadgeCount) {
      ctx.fillStyle = "red";
      ctx.beginPath();
      ctx.ellipse(
        BADGE_CENTRE_X,
        BADGE_CENTRE_Y,
        BADGE_RADIUS,
        BADGE_RADIUS,
        0,
        0,
        2 * Math.PI,
      );
      ctx.fill();

      ctx.textAlign = "center";
      ctx.fillStyle = "white";
      ctx.font =
        'bold 70px "Segoe UI","Helvetica Neue",Helvetica,Arial,sans-serif';
      ctx.fillText(
        count > 9 ? "+" : count.toString(),
        BADGE_CENTRE_X,
        BADGE_TEXT_BASELINE_Y,
      );
    }

    return this.#resizeToIconSize(canvas).toDataURL();
  }

  #resizeToIconSize(canvas) {
    const resized = document.createElement("canvas");
    resized.width = this.iconSize.width;
    resized.height = this.iconSize.height;

    const ctx = resized.getContext("2d");
    ctx.scale(
      this.iconSize.width / canvas.width,
      this.iconSize.height / canvas.height,
    );
    ctx.drawImage(canvas, 0, 0);

    return resized;
  }
}

module.exports = new TrayIconRenderer();
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/trayIconRenderer.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/browser/tools/trayIconRenderer.js tests/unit/trayIconRenderer.test.js
git commit -m "feat: render the unread badge onto the tray icon"
```

---

### Task 13: Application tray

**Files:**
- Create: `app/menus/tray.js`
- Test: `tests/unit/tray.test.js`

**Interfaces:**
- Consumes: `config.appTitle`, `config.trayIconEnabled` (Task 3); the `tray-update` payload shape from Task 12.
- Produces: `module.exports = ApplicationTray`.
  - `new ApplicationTray(window, config, iconPath, deps = {})` where `deps` supplies `{ Tray, Menu, nativeImage }`, defaulting to `require("electron")`. `window` is a `BrowserWindow`-like object.
  - `initialize(ipcMain) => void` — creates the tray and registers `ipcMain.on("tray-update", ...)`.
  - `updateTrayImage({ icon, flash, count }) => void` — `icon` is a data URL or `null` (fall back to the base icon path).
  - `close() => void` — destroys the tray.
  - Consumed by `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/tray.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const ApplicationTray = require("../../app/menus/tray");

function fakeDeps() {
  const created = [];
  class FakeTray {
    constructor(image) {
      this.image = image;
      this.tooltip = null;
      this.menu = null;
      this.destroyed = false;
      this.listeners = {};
      created.push(this);
    }
    setToolTip(text) {
      this.tooltip = text;
    }
    setContextMenu(menu) {
      this.menu = menu;
    }
    setImage(image) {
      this.image = image;
    }
    on(event, handler) {
      this.listeners[event] = handler;
    }
    destroy() {
      this.destroyed = true;
    }
  }
  return {
    created,
    deps: {
      Tray: FakeTray,
      Menu: { buildFromTemplate: (template) => ({ template }) },
      nativeImage: {
        createFromDataURL: (url) => ({ from: "dataURL", url }),
        createFromPath: (p) => ({ from: "path", path: p }),
      },
    },
  };
}

function fakeWindow() {
  return {
    visible: true,
    flashed: null,
    focused: false,
    isVisible() {
      return this.visible;
    },
    show() {
      this.visible = true;
    },
    hide() {
      this.visible = false;
    },
    focus() {
      this.focused = true;
    },
    flashFrame(value) {
      this.flashed = value;
    },
  };
}

function fakeIpcMain() {
  const handlers = {};
  return {
    handlers,
    on(channel, handler) {
      handlers[channel] = handler;
    },
  };
}

const CONFIG = { appTitle: "Microsoft Outlook", trayIconEnabled: true };

test("creates a tray with the base icon, tooltip and a context menu", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(fakeWindow(), CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  assert.strictEqual(created.length, 1);
  assert.deepStrictEqual(created[0].image, { from: "path", path: "/icons/base.png" });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook");
  assert.ok(created[0].menu, "a context menu must be attached");
});

test("does not create a tray when the tray icon is disabled", () => {
  const { created, deps } = fakeDeps();
  const ipcMain = fakeIpcMain();
  const tray = new ApplicationTray(
    fakeWindow(),
    { ...CONFIG, trayIconEnabled: false },
    "/icons/base.png",
    deps,
  );
  tray.initialize(ipcMain);

  assert.strictEqual(created.length, 0);
  assert.strictEqual(ipcMain.handlers["tray-update"], undefined);
});

test("a tray-update with a count swaps the image and flashes the window", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  const ipcMain = fakeIpcMain();
  tray.initialize(ipcMain);

  ipcMain.handlers["tray-update"]({}, {
    icon: "data:image/png;base64,AAAA",
    flash: true,
    count: 4,
  });

  assert.deepStrictEqual(created[0].image, {
    from: "dataURL",
    url: "data:image/png;base64,AAAA",
  });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook (4)");
  assert.strictEqual(window.flashed, true);
});

test("a zero count restores the base icon and plain tooltip", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  tray.updateTrayImage({ icon: null, flash: false, count: 0 });

  assert.deepStrictEqual(created[0].image, { from: "path", path: "/icons/base.png" });
  assert.strictEqual(created[0].tooltip, "Microsoft Outlook");
  assert.strictEqual(window.flashed, false);
});

test("clicking the tray shows and focuses a hidden window", () => {
  const { created, deps } = fakeDeps();
  const window = fakeWindow();
  window.visible = false;
  const tray = new ApplicationTray(window, CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  created[0].listeners.click();

  assert.strictEqual(window.visible, true);
  assert.strictEqual(window.focused, true);
});

test("a malformed tray-update payload does not throw", () => {
  const { deps } = fakeDeps();
  const tray = new ApplicationTray(fakeWindow(), CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  assert.doesNotThrow(() => tray.updateTrayImage(undefined));
  assert.doesNotThrow(() => tray.updateTrayImage({}));
});

test("close destroys the tray and is safe to call twice", () => {
  const { created, deps } = fakeDeps();
  const tray = new ApplicationTray(fakeWindow(), CONFIG, "/icons/base.png", deps);
  tray.initialize(fakeIpcMain());

  tray.close();
  tray.close();

  assert.strictEqual(created[0].destroyed, true);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/tray.test.js`
Expected: FAIL — `Cannot find module '../../app/menus/tray'`.

- [ ] **Step 3: Create `app/menus/tray.js`**

```javascript
/**
 * Owns the system tray icon and its context menu.
 *
 * Electron collaborators are injected so this module can be unit tested
 * outside a running Electron process.
 */
class ApplicationTray {
  #window;
  #config;
  #iconPath;
  #Tray;
  #Menu;
  #nativeImage;
  #tray = null;

  /**
   * @param {object} window - The main BrowserWindow.
   * @param {object} config - The resolved startup config.
   * @param {string} iconPath - Absolute path to the base tray icon.
   * @param {object} [deps] - Electron collaborators, injectable for tests.
   */
  constructor(window, config, iconPath, deps = {}) {
    const electron = deps.Tray ? deps : require("electron");
    this.#window = window;
    this.#config = config;
    this.#iconPath = iconPath;
    this.#Tray = electron.Tray;
    this.#Menu = electron.Menu;
    this.#nativeImage = electron.nativeImage;
  }

  /** @param {object} ipcMain - Electron's ipcMain. */
  initialize(ipcMain) {
    if (!this.#config.trayIconEnabled) {
      console.info("[TRAY] Tray icon disabled by config");
      return;
    }

    try {
      this.#tray = new this.#Tray(this.#nativeImage.createFromPath(this.#iconPath));
      this.#tray.setToolTip(this.#config.appTitle);
      this.#tray.setContextMenu(this.#Menu.buildFromTemplate(this.#buildMenu()));
      this.#tray.on("click", () => this.#toggleWindow());
      console.info("[TRAY] Tray created");
    } catch (error) {
      console.error("[TRAY] Failed to create tray:", error.message);
      return;
    }

    ipcMain.on("tray-update", (_event, payload) => this.updateTrayImage(payload));
  }

  #buildMenu() {
    return [
      { label: "Show / Hide", click: () => this.#toggleWindow() },
      { label: "Reload", click: () => this.#window.reload?.() },
      { type: "separator" },
      { label: "Quit", click: () => this.#window.emit?.("tray-quit") },
    ];
  }

  #toggleWindow() {
    if (this.#window.isVisible()) {
      this.#window.hide();
      return;
    }
    this.#window.show();
    this.#window.focus();
  }

  /**
   * @param {{icon: string|null, flash: boolean, count: number}} payload
   */
  updateTrayImage(payload) {
    if (!this.#tray) {
      return;
    }

    try {
      const { icon = null, flash = false, count = 0 } = payload ?? {};

      // A null icon means "no badge": fall back to the icon on disk rather
      // than leaving a stale badged image in the tray.
      this.#tray.setImage(
        icon
          ? this.#nativeImage.createFromDataURL(icon)
          : this.#nativeImage.createFromPath(this.#iconPath),
      );
      this.#tray.setToolTip(
        count > 0 ? `${this.#config.appTitle} (${count})` : this.#config.appTitle,
      );
      this.#window.flashFrame(Boolean(flash));
    } catch (error) {
      console.error("[TRAY] Failed to update tray image:", error.message);
    }
  }

  close() {
    if (!this.#tray) {
      return;
    }
    this.#tray.destroy();
    this.#tray = null;
  }
}

module.exports = ApplicationTray;
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/tray.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/menus/tray.js tests/unit/tray.test.js
git commit -m "feat: add system tray with unread badge updates"
```

---

### Task 14: Native notification service

Receives the `show-notification` IPC that the preload Notification shim (Task 15) sends, and renders it as a native OS notification.

**Files:**
- Create: `app/notifications/service.js`
- Test: `tests/unit/notificationService.test.js`

**Interfaces:**
- Consumes: `config.disableNotifications`, `config.defaultNotificationUrgency` (Task 3).
- Produces: `module.exports = NotificationService`.
  - `new NotificationService(config, deps = {})` where `deps` supplies `{ Notification, nativeImage }`, defaulting to `require("electron")`.
  - `setWindow(window) => void` — the window to focus on click.
  - `initialize(ipcMain) => void` — registers `ipcMain.handle("show-notification", ...)`.
  - `show(options) => Promise<void>` where `options` is `{ id?: string|number, title?: string, body?: string, icon?: string, urgency?: string }`. Never rejects.
  - Sends `notification-closed` with `{ id }` to the renderer when a notification is clicked or dismissed.

> **Simplification vs teams-for-linux:** teams-for-linux also fetches remote same-origin avatar icons over https with a timeout and a byte cap. This MVP accepts only `data:` URL icons and drops anything else — no network fetch, no timeout budget, no size cap to get wrong. Remote avatars can be added later behind their own spec.
>
> **PII:** notification titles and bodies contain sender names, addresses and mail subjects. Log only that a notification was shown, never its content.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/notificationService.test.js`:

```javascript
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
        createFromDataURL: (url) => ({ from: "dataURL", url, isEmpty: () => false }),
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

const CONFIG = { disableNotifications: false, defaultNotificationUrgency: "normal" };

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
  const service = new NotificationService({ ...CONFIG, disableNotifications: true }, deps);
  await service.show({ title: "New mail" });
  assert.strictEqual(created.length, 0);
});

test("accepts a data URL icon and ignores any other icon source", async () => {
  const { created, deps } = fakeDeps();
  const service = new NotificationService(CONFIG, deps);

  await service.show({ title: "a", icon: "data:image/png;base64,AAAA" });
  assert.strictEqual(created[0].options.icon.from, "dataURL");
  assert.strictEqual(created[0].options.icon.url, "data:image/png;base64,AAAA");

  await service.show({ title: "b", icon: "https://outlook.office.com/avatar.png" });
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
  service.initialize({ handle: (channel, handler) => (handlers[channel] = handler) });

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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/notificationService.test.js`
Expected: FAIL — `Cannot find module '../../app/notifications/service'`.

- [ ] **Step 3: Create `app/notifications/service.js`**

```javascript
const DEFAULT_TITLE = "Microsoft Outlook";
const VALID_URGENCIES = new Set(["low", "normal", "critical"]);

/**
 * Renders native OS notifications on behalf of the Outlook page.
 *
 * The page's own `Notification` calls are intercepted in preload and arrive
 * here over IPC, so notifications survive the page being backgrounded and
 * respect the user's desktop notification settings.
 */
class NotificationService {
  #config;
  #Notification;
  #nativeImage;
  #window = null;

  /**
   * @param {object} config - The resolved startup config.
   * @param {object} [deps] - Electron collaborators, injectable for tests.
   */
  constructor(config, deps = {}) {
    const electron = deps.Notification ? deps : require("electron");
    this.#config = config;
    this.#Notification = electron.Notification;
    this.#nativeImage = electron.nativeImage;
  }

  /** @param {object} window - The main BrowserWindow. */
  setWindow(window) {
    this.#window = window;
  }

  /** @param {object} ipcMain - Electron's ipcMain. */
  initialize(ipcMain) {
    // Renders a native OS notification forwarded from the Outlook page.
    ipcMain.handle("show-notification", (_event, options) => this.show(options));
  }

  /**
   * @param {object} options - `{ id, title, body, icon, urgency }`.
   * @returns {Promise<void>} Always resolves.
   */
  async show(options) {
    if (this.#config.disableNotifications) {
      return;
    }

    const { id, title, body, icon, urgency } = options ?? {};

    try {
      const notification = new this.#Notification({
        title: typeof title === "string" && title !== "" ? title : DEFAULT_TITLE,
        body: typeof body === "string" ? body : "",
        icon: this.#loadIcon(icon),
        urgency: VALID_URGENCIES.has(urgency)
          ? urgency
          : this.#config.defaultNotificationUrgency,
      });

      // Both paths report back so the page can clear its own state; guard so a
      // click followed by the platform's close event only reports once.
      let reported = false;
      const report = () => {
        if (reported) {
          return;
        }
        reported = true;
        this.#window?.webContents.send("notification-closed", { id });
      };

      notification.on("click", () => {
        this.#focusWindow();
        report();
      });
      notification.on("close", report);

      notification.show();
      // Never log title or body: they carry sender names and mail subjects.
      console.debug("[NOTIFY] Notification shown");
    } catch (error) {
      console.error("[NOTIFY] Failed to show notification:", error.message);
    }
  }

  /**
   * Only `data:` URLs are supported. Anything else (including remote avatar
   * URLs) is dropped so this path never makes a network request.
   */
  #loadIcon(icon) {
    if (typeof icon !== "string" || !icon.startsWith("data:")) {
      return undefined;
    }
    try {
      const image = this.#nativeImage.createFromDataURL(icon);
      return image.isEmpty() ? undefined : image;
    } catch (error) {
      console.error("[NOTIFY] Failed to decode notification icon:", error.message);
      return undefined;
    }
  }

  #focusWindow() {
    try {
      this.#window?.show();
      this.#window?.focus();
    } catch (error) {
      console.error("[NOTIFY] Failed to focus window:", error.message);
    }
  }
}

module.exports = NotificationService;
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/notificationService.test.js`
Expected: PASS — 10 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/notifications/service.js tests/unit/notificationService.test.js
git commit -m "feat: forward Outlook notifications to native OS notifications"
```

---

### Task 15: Preload bridge and Notification shim

The hinge of the whole app: the shim replaces the page's `Notification` constructor so Outlook's new-mail notifications become IPC messages instead of Chromium notifications. The shim is a separate, tested module; `preload.js` is thin wiring around it.

**Files:**
- Create: `app/browser/tools/notificationShim.js`
- Create: `app/browser/preload.js`
- Test: `tests/unit/notificationShim.test.js`

**Interfaces:**
- Consumes: `ipcRenderer`; `mutationTitle` (Task 11); `trayIconRenderer` (Task 12); the `show-notification` handler (Task 14); the `notification-closed` send (Task 14).
- Produces: `module.exports = { installNotificationShim }`.
  - `installNotificationShim(target: object, ipcRenderer: object) => void` — replaces `target.Notification`. `target` is `globalThis` in production, a plain object in tests.
  - The installed class: `new Notification(title, options)` invokes `show-notification` with `{ id, title, body, icon, urgency }`; `static requestPermission()` resolves `"granted"`; `static get permission()` returns `"granted"`; instances support `onclick`, `onclose`, `close()`, and `addEventListener("click"|"close", fn)`.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/notificationShim.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const { installNotificationShim } = require("../../app/browser/tools/notificationShim");

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
  assert.doesNotThrow(() => channels["notification-closed"]({}, { id: 99999 }));
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/notificationShim.test.js`
Expected: FAIL — `Cannot find module '../../app/browser/tools/notificationShim'`.

- [ ] **Step 3: Create `app/browser/tools/notificationShim.js`**

```javascript
/**
 * Replaces the page's `Notification` constructor so Outlook's new-mail
 * notifications are rendered by the main process as native OS notifications
 * instead of by Chromium.
 *
 * Chromium's own notifications inside a wrapped web app do not respect the
 * desktop's notification settings and vanish when the page is backgrounded,
 * which is the behaviour users report as "notifications don't work".
 */
function installNotificationShim(target, ipcRenderer) {
  const live = new Map();
  let nextId = 1;

  class ShimNotification {
    #id;
    #listeners = { click: [], close: [] };
    #closed = false;

    constructor(title, options = {}) {
      this.#id = nextId++;
      this.title = typeof title === "string" ? title : "";
      this.body = typeof options.body === "string" ? options.body : "";
      this.icon = options.icon;
      this.onclick = null;
      this.onclose = null;

      live.set(this.#id, this);

      // Fire-and-forget: the page's Notification constructor is synchronous,
      // so a rejection here has nowhere to go but a log line.
      Promise.resolve(
        ipcRenderer.invoke("show-notification", {
          id: this.#id,
          title: this.title,
          body: this.body,
          icon: this.icon,
          urgency: options.urgency,
        }),
      ).catch((error) => {
        // Never log title or body: they carry sender names and mail subjects.
        console.error("[NOTIFY] Failed to forward notification:", error.message);
        live.delete(this.#id);
      });
    }

    get id() {
      return this.#id;
    }

    addEventListener(type, handler) {
      if (this.#listeners[type] && typeof handler === "function") {
        this.#listeners[type].push(handler);
      }
    }

    removeEventListener(type, handler) {
      const handlers = this.#listeners[type];
      if (!handlers) {
        return;
      }
      const index = handlers.indexOf(handler);
      if (index !== -1) {
        handlers.splice(index, 1);
      }
    }

    close() {
      this.#dispatch("close");
    }

    /** Called from the main-process reply. */
    handleRemoteClose() {
      this.#dispatch("close");
    }

    #dispatch(type) {
      if (this.#closed) {
        return;
      }
      this.#closed = true;
      live.delete(this.#id);

      try {
        this[`on${type}`]?.call(this, { type, target: this });
        for (const handler of this.#listeners[type]) {
          handler.call(this, { type, target: this });
        }
      } catch (error) {
        console.error(`[NOTIFY] ${type} handler threw:`, error.message);
      }
    }

    static requestPermission(callback) {
      // The user already granted permission by installing a mail client.
      if (typeof callback === "function") {
        callback("granted");
      }
      return Promise.resolve("granted");
    }

    static get permission() {
      return "granted";
    }
  }

  ipcRenderer.on("notification-closed", (_event, payload) => {
    live.get(payload?.id)?.handleRemoteClose();
  });

  target.Notification = ShimNotification;
}

module.exports = { installNotificationShim };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/notificationShim.test.js`
Expected: PASS — 10 tests.

- [ ] **Step 5: Create `app/browser/preload.js`**

Thin wiring, verified by running the app in Task 18 rather than by unit test — there is no logic here to assert that is not already covered by the modules it calls.

```javascript
const { ipcRenderer } = require("electron");
const { installNotificationShim } = require("./tools/notificationShim");
const mutationTitle = require("./tools/mutationTitle");
const trayIconRenderer = require("./tools/trayIconRenderer");

// Installed synchronously, before any page script runs. Waiting for the
// config round-trip would let Outlook capture the real Notification
// constructor first, and notifications would bypass the shim entirely.
installNotificationShim(globalThis, ipcRenderer);

// Surface renderer-side failures in the main log, where users find them.
globalThis.addEventListener("error", (event) => {
  ipcRenderer.send("window-error", {
    message: event.message,
    source: event.filename,
    line: event.lineno,
  });
});

globalThis.addEventListener("unhandledrejection", (event) => {
  ipcRenderer.send("unhandled-rejection", {
    reason: String(event.reason?.message ?? event.reason),
  });
});

globalThis.addEventListener("DOMContentLoaded", async () => {
  try {
    const config = await ipcRenderer.invoke("get-config");

    // Both are config-gated and DOM-dependent, so they wait for the config
    // round-trip; the notification shim above deliberately does not.
    mutationTitle.init(config);
    if (config.trayIconEnabled) {
      trayIconRenderer.init(config, ipcRenderer);
    }
    console.debug("[PRELOAD] Browser tools initialised");
  } catch (error) {
    console.error("[PRELOAD] Initialisation failed:", error.message);
  }
});
```

- [ ] **Step 6: Lint and commit**

```bash
npm run lint
npm run test:unit
git add app/browser/tools/notificationShim.js app/browser/preload.js tests/unit/notificationShim.test.js
git commit -m "feat: intercept page notifications and wire up browser tools"
```

---

### Task 16: Main application window

**Files:**
- Create: `app/mainAppWindow/index.js`
- Test: `tests/unit/mainAppWindow.test.js`

**Interfaces:**
- Consumes: `config.url`, `config.appTitle`, `config.closeToTray`, `config.trayIconEnabled`, `config.webDebug` (Task 3).
- Produces: `module.exports = { createWindow, shouldOpenExternally }`.
  - `createWindow(config: object, deps = {}) => object` — returns the created window. `deps` supplies `{ BrowserWindow, shell, preloadPath }`, defaulting to `require("electron")` and the path to `../browser/preload.js`.
  - `shouldOpenExternally(url: string, appUrl: string) => boolean` — true when a URL should go to the system browser instead of the app window. Exported for testing.
  - Consumed by `app/index.js` (Task 17).

- [ ] **Step 1: Write the failing test**

Create `tests/unit/mainAppWindow.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const { createWindow, shouldOpenExternally } = require("../../app/mainAppWindow");

function fakeDeps() {
  const opened = [];
  const created = [];
  class FakeBrowserWindow {
    constructor(options) {
      this.options = options;
      this.loadedUrl = null;
      this.listeners = {};
      this.hidden = false;
      this.webContents = {
        handlers: {},
        setWindowOpenHandler(handler) {
          this.handlers.windowOpen = handler;
        },
        openDevTools() {
          this.devToolsOpen = true;
        },
      };
      created.push(this);
    }
    loadURL(url) {
      this.loadedUrl = url;
    }
    on(event, handler) {
      this.listeners[event] = handler;
    }
    hide() {
      this.hidden = true;
    }
  }
  return {
    created,
    opened,
    deps: {
      BrowserWindow: FakeBrowserWindow,
      shell: { openExternal: (url) => opened.push(url) },
      preloadPath: "/app/browser/preload.js",
    },
  };
}

const CONFIG = {
  url: "https://outlook.office.com",
  appTitle: "Microsoft Outlook",
  closeToTray: true,
  trayIconEnabled: true,
  webDebug: false,
};

test("loads the configured Outlook URL with the preload script attached", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);

  assert.strictEqual(created[0].loadedUrl, "https://outlook.office.com");
  assert.strictEqual(created[0].options.webPreferences.preload, "/app/browser/preload.js");
  assert.strictEqual(created[0].options.title, "Microsoft Outlook");
});

test("respects a custom URL from config", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, url: "https://outlook.live.com/mail/0/" }, deps);
  assert.strictEqual(created[0].loadedUrl, "https://outlook.live.com/mail/0/");
});

test("opens dev tools only when web debugging is enabled", () => {
  const off = fakeDeps();
  createWindow(CONFIG, off.deps);
  assert.notStrictEqual(off.created[0].webContents.devToolsOpen, true);

  const on = fakeDeps();
  createWindow({ ...CONFIG, webDebug: true }, on.deps);
  assert.strictEqual(on.created[0].webContents.devToolsOpen, true);
});

test("closing hides the window when closeToTray is on", () => {
  const { created, deps } = fakeDeps();
  createWindow(CONFIG, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, true);
  assert.strictEqual(created[0].hidden, true);
});

test("closing really closes when closeToTray is off", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, closeToTray: false }, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, false);
  assert.strictEqual(created[0].hidden, false);
});

test("closeToTray is ignored when there is no tray to close to", () => {
  const { created, deps } = fakeDeps();
  createWindow({ ...CONFIG, trayIconEnabled: false }, deps);

  let prevented = false;
  created[0].listeners.close({ preventDefault: () => (prevented = true) });

  assert.strictEqual(prevented, false, "hiding with no tray would trap the window");
});

test("third-party links open in the system browser, Outlook links stay in-app", () => {
  const { created, deps, opened } = fakeDeps();
  createWindow(CONFIG, deps);
  const handler = created[0].webContents.handlers.windowOpen;

  assert.deepStrictEqual(handler({ url: "https://example.com/report" }), {
    action: "deny",
  });
  assert.deepStrictEqual(opened, ["https://example.com/report"]);

  assert.deepStrictEqual(
    handler({ url: "https://outlook.office.com/mail/inbox/id/xyz" }),
    { action: "allow" },
  );
});

test("Microsoft sign-in URLs stay in-app so authentication can complete", () => {
  assert.strictEqual(
    shouldOpenExternally(
      "https://login.microsoftonline.com/common/oauth2/authorize",
      "https://outlook.office.com",
    ),
    false,
  );
  assert.strictEqual(
    shouldOpenExternally("https://outlook.office365.com/mail", "https://outlook.office.com"),
    false,
  );
});

test("non-http schemes are handed to the OS", () => {
  assert.strictEqual(
    shouldOpenExternally("mailto:someone@example.com", "https://outlook.office.com"),
    true,
  );
});

test("an unparseable URL is not opened externally", () => {
  assert.strictEqual(shouldOpenExternally("not a url", "https://outlook.office.com"), false);
  assert.strictEqual(shouldOpenExternally(undefined, "https://outlook.office.com"), false);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/mainAppWindow.test.js`
Expected: FAIL — `Cannot find module '../../app/mainAppWindow'`.

- [ ] **Step 3: Create `app/mainAppWindow/index.js`**

```javascript
const path = require("node:path");

const DEFAULT_WIDTH = 1280;
const DEFAULT_HEIGHT = 800;

// Hosts that are part of the Outlook experience or its sign-in flow. Sending
// any of these to the system browser would break authentication.
const IN_APP_HOST_SUFFIXES = [
  "outlook.office.com",
  "outlook.office365.com",
  "outlook.live.com",
  "office.com",
  "login.microsoftonline.com",
  "login.live.com",
  "login.microsoft.com",
  "microsoftonline.com",
];

/**
 * Decides whether a navigation target belongs in the system browser.
 *
 * @param {string} url - The target URL.
 * @param {string} appUrl - The configured Outlook URL.
 * @returns {boolean} True when the OS should handle it.
 */
function shouldOpenExternally(url, appUrl) {
  let target;
  try {
    target = new URL(url);
  } catch {
    // Unparseable input is never handed to the shell.
    return false;
  }

  if (target.protocol !== "http:" && target.protocol !== "https:") {
    // mailto:, tel: and friends are the OS's business.
    return true;
  }

  let appHost = "";
  try {
    appHost = new URL(appUrl).hostname;
  } catch {
    appHost = "";
  }

  const host = target.hostname;
  const isInApp =
    host === appHost ||
    IN_APP_HOST_SUFFIXES.some(
      (suffix) => host === suffix || host.endsWith(`.${suffix}`),
    );

  return !isInApp;
}

/**
 * @param {object} config - The resolved startup config.
 * @param {object} [deps] - Electron collaborators, injectable for tests.
 * @returns {object} The created window.
 */
function createWindow(config, deps = {}) {
  const electron = deps.BrowserWindow ? deps : require("electron");
  const { BrowserWindow, shell } = electron;
  const preload =
    deps.preloadPath ?? path.join(__dirname, "..", "browser", "preload.js");

  const window = new BrowserWindow({
    width: DEFAULT_WIDTH,
    height: DEFAULT_HEIGHT,
    title: config.appTitle,
    autoHideMenuBar: true,
    webPreferences: {
      preload,
      // The preload script needs `require` to reach Electron's nativeImage for
      // tray badge rendering, so it is not sandboxed and context isolation is
      // off. The IPC allowlist in app/security/ipcValidator.js is the
      // compensating control: the page can only reach the channels listed
      // there, with payloads sanitised.
      sandbox: false,
      contextIsolation: false,
      nodeIntegration: false,
      spellcheck: false,
    },
  });

  window.on("close", (event) => {
    // Only hide when there is a tray to restore from; hiding without one
    // would leave the app running with no way to get it back.
    if (config.closeToTray && config.trayIconEnabled) {
      event.preventDefault();
      window.hide();
    }
  });

  window.webContents.setWindowOpenHandler(({ url }) => {
    if (shouldOpenExternally(url, config.url)) {
      shell.openExternal(url);
      return { action: "deny" };
    }
    return { action: "allow" };
  });

  if (config.webDebug) {
    window.webContents.openDevTools();
  }

  window.loadURL(config.url);
  console.info("[WINDOW] Main window created");

  return window;
}

module.exports = { createWindow, shouldOpenExternally };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/mainAppWindow.test.js`
Expected: PASS — 10 tests.

- [ ] **Step 5: Lint and commit**

```bash
npm run lint
git add app/mainAppWindow/index.js tests/unit/mainAppWindow.test.js
git commit -m "feat: create the main window that loads Outlook web"
```

---

### Task 17: IPC registration and main process entry

The last wiring task. `app/ipc/register.js` holds the testable part — every IPC registration passes through the allowlist; `app/index.js` is the assembly root.

**Files:**
- Create: `app/ipc/register.js`
- Create: `app/index.js`
- Test: `tests/unit/ipcRegister.test.js`

**Interfaces:**
- Consumes: `validateIpcChannel` (Task 8); `argv` (Task 6); `AppConfiguration` (Task 7); `isNetworkError` (Task 4); `CommandLineManager.applySwitches` (Task 9); `TrayIconChooser` (Task 10); `createWindow` (Task 16); `ApplicationTray` (Task 13); `NotificationService` (Task 14).
- Produces: `app/ipc/register.js` → `module.exports = { createGuardedIpcMain, registerIpcHandlers }`.
  - `createGuardedIpcMain(ipcMain: object) => {on, handle}` — a façade that refuses to register a channel missing from the allowlist (throwing at startup, not silently at runtime) and validates every inbound payload before the handler sees it.
  - `registerIpcHandlers(guardedIpcMain: object, deps: {config: object, setBadgeCount: (n: number) => void}) => void` — registers `get-config`, `set-badge-count`, `window-error`, `unhandled-rejection`.

- [ ] **Step 1: Write the failing test**

Create `tests/unit/ipcRegister.test.js`:

```javascript
const test = require("node:test");
const assert = require("node:assert");
const {
  createGuardedIpcMain,
  registerIpcHandlers,
} = require("../../app/ipc/register");

function fakeIpcMain() {
  const on = {};
  const handle = {};
  return {
    on: (channel, handler) => (on[channel] = handler),
    handle: (channel, handler) => (handle[channel] = handler),
    registered: { on, handle },
  };
}

test("registering a channel that is not allowlisted throws at startup", () => {
  const guarded = createGuardedIpcMain(fakeIpcMain());
  assert.throws(
    () => guarded.handle("definitely-not-allowed", () => {}),
    /not allowlisted/i,
  );
});

test("an allowlisted channel reaches the real ipcMain", () => {
  const ipcMain = fakeIpcMain();
  const guarded = createGuardedIpcMain(ipcMain);
  guarded.handle("get-config", () => "ok");
  guarded.on("tray-update", () => {});

  assert.strictEqual(typeof ipcMain.registered.handle["get-config"], "function");
  assert.strictEqual(typeof ipcMain.registered.on["tray-update"], "function");
});

test("payloads are sanitized before the handler sees them", () => {
  const ipcMain = fakeIpcMain();
  const guarded = createGuardedIpcMain(ipcMain);

  let seen;
  guarded.on("tray-update", (_event, payload) => (seen = payload));

  const payload = JSON.parse('{"count":3,"__proto__":{"polluted":true}}');
  ipcMain.registered.on["tray-update"]({}, payload);

  assert.strictEqual(seen.count, 3);
  assert.strictEqual({}.polluted, undefined);
});

test("get-config returns the resolved config", async () => {
  const ipcMain = fakeIpcMain();
  const config = { url: "https://outlook.office.com", trayIconEnabled: true };
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config,
    setBadgeCount: () => {},
  });

  assert.deepStrictEqual(await ipcMain.registered.handle["get-config"]({}), config);
});

test("set-badge-count forwards a valid count and rejects nonsense", async () => {
  const ipcMain = fakeIpcMain();
  const counts = [];
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: (n) => counts.push(n),
  });

  const handler = ipcMain.registered.handle["set-badge-count"];
  await handler({}, 4);
  await handler({}, 0);
  await handler({}, "many");
  await handler({}, -1);

  assert.deepStrictEqual(counts, [4, 0]);
});

test("a failing setBadgeCount does not reject the invoke", async () => {
  const ipcMain = fakeIpcMain();
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: () => {
      throw new Error("unsupported on this desktop");
    },
  });

  await assert.doesNotReject(() =>
    ipcMain.registered.handle["set-badge-count"]({}, 2),
  );
});

test("renderer error channels are registered and do not throw", () => {
  const ipcMain = fakeIpcMain();
  registerIpcHandlers(createGuardedIpcMain(ipcMain), {
    config: {},
    setBadgeCount: () => {},
  });

  assert.doesNotThrow(() =>
    ipcMain.registered.on["window-error"]({}, { message: "boom", line: 12 }),
  );
  assert.doesNotThrow(() =>
    ipcMain.registered.on["unhandled-rejection"]({}, { reason: "boom" }),
  );
  assert.doesNotThrow(() => ipcMain.registered.on["window-error"]({}, undefined));
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node --test tests/unit/ipcRegister.test.js`
Expected: FAIL — `Cannot find module '../../app/ipc/register'`.

- [ ] **Step 3: Create `app/ipc/register.js`**

```javascript
const { validateIpcChannel, allowedChannels } = require("../security/ipcValidator");

const MAX_BADGE_COUNT = 9999;

/**
 * Wraps ipcMain so no channel can be registered or served without passing the
 * allowlist.
 *
 * Registering an unlisted channel throws immediately: a missing allowlist
 * entry is a programming error, and failing at startup is far easier to
 * diagnose than a handler that silently never fires.
 *
 * @param {object} ipcMain - Electron's ipcMain.
 * @returns {{on: Function, handle: Function}} The guarded façade.
 */
function createGuardedIpcMain(ipcMain) {
  const assertAllowed = (channel) => {
    if (!allowedChannels.has(channel)) {
      throw new Error(
        `IPC channel "${channel}" is not allowlisted. Add it to app/security/ipcValidator.js.`,
      );
    }
  };

  const guard = (channel, handler) => (event, payload) => {
    if (!validateIpcChannel(channel, payload)) {
      return undefined;
    }
    return handler(event, payload);
  };

  return {
    on(channel, handler) {
      assertAllowed(channel);
      ipcMain.on(channel, guard(channel, handler));
    },
    handle(channel, handler) {
      assertAllowed(channel);
      const guarded = guard(channel, handler);
      // `invoke()` in the renderer always yields a promise, so the registered
      // handler presents one too, whether or not it is itself async. A
      // synchronous handler that returned a bare value would still work in
      // Electron, but callers here can rely on the promise contract.
      ipcMain.handle(channel, async (event, payload) =>
        guarded(event, payload),
      );
    },
  };
}

/**
 * Registers the channels that belong to the main process itself. Tray and
 * notification channels are registered by their own modules, through the same
 * guarded façade.
 *
 * @param {object} guardedIpcMain - From createGuardedIpcMain.
 * @param {object} deps - `{ config, setBadgeCount }`.
 */
function registerIpcHandlers(guardedIpcMain, { config, setBadgeCount }) {
  // Returns the resolved startup config to the renderer.
  guardedIpcMain.handle("get-config", () => config);

  // Sets the OS taskbar/dock badge to the unread count.
  guardedIpcMain.handle("set-badge-count", (_event, count) => {
    if (!Number.isInteger(count) || count < 0 || count > MAX_BADGE_COUNT) {
      console.warn("[IPC] Ignoring out-of-range badge count");
      return;
    }
    try {
      setBadgeCount(count);
    } catch (error) {
      // Not every desktop environment supports badges; a failure here must
      // not surface as a rejected invoke in the page.
      console.debug("[IPC] Badge count not applied:", error.message);
    }
  });

  // Renderer-side uncaught error, forwarded from preload.
  guardedIpcMain.on("window-error", (_event, payload) => {
    console.error(
      `[RENDERER] ${payload?.message ?? "unknown error"} (${payload?.source ?? "?"}:${payload?.line ?? "?"})`,
    );
  });

  // Renderer-side unhandled promise rejection, forwarded from preload.
  guardedIpcMain.on("unhandled-rejection", (_event, payload) => {
    console.error(`[RENDERER] Unhandled rejection: ${payload?.reason ?? "unknown"}`);
  });
}

module.exports = { createGuardedIpcMain, registerIpcHandlers };
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node --test tests/unit/ipcRegister.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Create `app/index.js`**

Assembly only — every behaviour it composes is tested in its own module.

```javascript
const { app, ipcMain } = require("electron");
const fs = require("node:fs");
const path = require("node:path");

const { isNetworkError } = require("./utils/networkErrors");
const { AppConfiguration } = require("./appConfiguration");
const CommandLineManager = require("./startup/commandLine");
const TrayIconChooser = require("./browser/tools/trayIconChooser");
const ApplicationTray = require("./menus/tray");
const NotificationService = require("./notifications/service");
const { createWindow } = require("./mainAppWindow");
const { createGuardedIpcMain, registerIpcHandlers } = require("./ipc/register");

// Registered before anything else so a failure during startup is logged
// rather than lost. A transient network error must not take the app down:
// Electron surfaces those as uncaught exceptions from the renderer's own
// load failures, and killing the process would look like a random crash to
// the user.
process.on("uncaughtException", (error) => {
  if (isNetworkError(error.message)) {
    console.warn("[NETWORK] Recoverable network error:", error.message);
    return;
  }
  console.error("[FATAL] Uncaught exception:", error.message, error.stack);
  process.exit(1);
});

process.on("unhandledRejection", (reason) => {
  const message = reason?.message ?? String(reason);
  if (isNetworkError(message)) {
    console.warn("[NETWORK] Recoverable network rejection:", message);
    return;
  }
  console.error("[FATAL] Unhandled rejection:", message);
  process.exit(1);
});

// A second instance would fight the first over the tray icon and the single
// Outlook session, so hand off to the running one instead.
if (!app.requestSingleInstanceLock()) {
  console.info("[STARTUP] Another instance is already running, exiting");
  app.quit();
  return;
}

const appConfiguration = new AppConfiguration(app.getPath("userData"), app.getVersion());
const config = appConfiguration.startupConfig;

CommandLineManager.applySwitches(config, app);

let window = null;
let tray = null;

app.on("second-instance", () => {
  window?.show();
  window?.focus();
});

app.on("window-all-closed", () => {
  tray?.close();
  app.quit();
});

app.whenReady().then(() => {
  const guardedIpcMain = createGuardedIpcMain(ipcMain);
  const notificationService = new NotificationService(config);

  registerIpcHandlers(guardedIpcMain, {
    config,
    setBadgeCount: (count) => app.setBadgeCount(count),
  });
  notificationService.initialize(guardedIpcMain);

  window = createWindow(config);
  notificationService.setWindow(window);

  tray = new ApplicationTray(
    window,
    config,
    new TrayIconChooser(config).getFile(),
  );
  tray.initialize(guardedIpcMain);

  // The tray's Quit item destroys rather than closes: closeToTray installs a
  // close handler that hides the window, which would otherwise make Quit a
  // no-op. destroy() skips that handler.
  window.on("tray-quit", () => {
    tray?.close();
    window?.destroy();
    app.quit();
  });

  if (config.watchConfigFile && config.isConfigFile) {
    watchConfigFile(appConfiguration.configPath);
  }

  console.info("[STARTUP] Application ready");
});

/**
 * Warns when the config file changes. Config is immutable after startup, so
 * this only tells the user a restart is needed — it never re-reads.
 */
function watchConfigFile(configPath) {
  try {
    fs.watch(path.join(configPath, "config.json"), { persistent: false }, () => {
      console.info("[CONFIG] Config file changed; restart to apply");
    });
  } catch (error) {
    console.error("[CONFIG] Failed to watch config file:", error.message);
  }
}
```

- [ ] **Step 6: Verify the app launches**

Run: `npm start`
Expected: a window opens showing the Microsoft sign-in page for Outlook, a tray icon appears, and the log shows `[STARTUP] Application ready` with no `[FATAL]` lines. Close the window: with the default `closeToTray: true` the app stays in the tray; quitting from the tray menu exits the process.

Then verify the config path is honoured:

```bash
mkdir -p ~/.config/outlook-for-linux
printf '{"appTitle":"Outlook (test)","trayIconEnabled":false}\n' > ~/.config/outlook-for-linux/config.json
npm start
```

Expected: the window title reads `Outlook (test)` and no tray icon appears. Remove the file afterwards.

> `app.getPath("userData")` resolves to `~/.config/outlook-for-linux` on Linux because the electron-builder `appId`/product name is `outlook-for-linux`, which is what the spec specifies as the config location.
>
> **Deviation found during execution:** in development this only holds if the
> start script is `electron .`, not `electron ./app`. Pointing Electron at
> `app/` makes that directory the app root; it has no `package.json`, so
> `app.getName()` falls back to `"Electron"` and the config is read from
> `~/.config/Electron` instead — the user's config file is silently ignored.
> The scripts in Task 1 use `electron .` for this reason.
>
> Note also that the OS window title tracks the page: `config.appTitle` is the
> initial title, and Outlook replaces it once loaded (that is the same title
> the unread-count observer reads). The config-file check above is therefore
> verified by `[CONFIG] Loaded user configuration` plus
> `[TRAY] Tray icon disabled by config` in the log, not by the title bar.

- [ ] **Step 7: Lint, test and commit**

```bash
npm run lint
npm run test:unit
git add app/ipc/register.js app/index.js tests/unit/ipcRegister.test.js
git commit -m "feat: wire up the main process entry point"
```

---

### Task 18: Packaging and documentation

**Files:**
- Modify: `README.md`
- Create: `docs/configuration.md`

**Interfaces:**
- Consumes: everything above.
- Produces: verified `AppImage`, `deb` and `rpm` artifacts in `dist/`, plus user-facing docs for every config option.

- [ ] **Step 1: Verify the unpacked build**

Run: `npm run pack`
Expected: exits 0 and produces `dist/linux-unpacked/outlook-for-linux`. Launch that binary directly and confirm the window opens — this catches assets missing from the package (the icon files in particular), which `npm start` cannot.

- [ ] **Step 2: Build the distributables**

Run: `npm run dist:linux`
Expected: exits 0. Confirm all three artifacts exist:

```bash
ls dist/*.AppImage dist/*.deb dist/*.rpm
```

If the rpm target fails for a missing `rpmbuild`, install it (`sudo dnf install rpm-build`) and re-run. Do not remove the rpm target to make the build pass — the spec requires it.

- [ ] **Step 3: Smoke-test the AppImage**

```bash
chmod +x dist/*.AppImage
./dist/*.AppImage
```

Expected: the window opens and the tray icon appears, exactly as with `npm start`.

- [ ] **Step 4: Write `docs/configuration.md`**

Document every option from `app/config/options.js` in a table: name, type, default, and what it does. Include the config file location (`~/.config/outlook-for-linux/config.json`), a worked example, and a note that CLI flags override the file which overrides the defaults, and that config is read once at startup.

```markdown
# Configuration

Configuration is read once at startup from
`~/.config/outlook-for-linux/config.json`. Command-line flags override the
file, and the file overrides the defaults. Changes require a restart.

Example:

```json
{
  "url": "https://outlook.live.com/mail/0/",
  "appTitle": "Outlook",
  "closeToTray": false,
  "defaultNotificationUrgency": "critical"
}
```

The same options work as flags: `outlook-for-linux --closeToTray=false`.

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `url` | string | `https://outlook.office.com` | Outlook endpoint to load. Use `https://outlook.live.com/mail/0/` for personal accounts, or your tenant URL. |
```

The table needs exactly one row per option in `app/config/options.js`, in this
order, with the Type and Default columns taken from that file and the
Description column copied verbatim from each option's `describe` string so the
two cannot drift:

`url`, `appTitle`, `appIcon`, `appIconType`, `trayIconEnabled`, `closeToTray`,
`disableNotifications`, `defaultNotificationUrgency`, `useMutationTitleLogic`,
`disableBadgeCount`, `disableNotificationWindowFlash`, `disableGpu`,
`logConfig`, `watchConfigFile`, `webDebug` — fifteen rows.

For `appIconType` and `defaultNotificationUrgency`, list the permitted
`choices` in the Description column. For `logConfig`, note that it is passed
through to `electron-log`, that `"console"` leaves the global console
untouched, and that `false` silences logging entirely.

Verify the count before committing:

```bash
grep -c '^  [a-zA-Z]*:' app/config/options.js   # expect 15
grep -c '^| `' docs/configuration.md            # expect 15
```

- [ ] **Step 5: Update `README.md`**

Replace the scaffolding placeholder with: what the project is (an unofficial Electron wrapper for Outlook web on Linux — state plainly that it is not affiliated with Microsoft), install instructions for each of the three artifacts, a link to `docs/configuration.md`, the development commands (`npm start`, `npm run lint`, `npm run test:unit`, `npm run dist:linux`), the GPL-3.0-or-later license line, and a short "not in this release" list taken from the spec's out-of-scope section so users are not surprised.

- [ ] **Step 6: Final full verification**

```bash
npm run lint
npm run test:unit
```

Expected: lint clean; all unit tests pass across every `tests/unit/*.test.js` file.

- [ ] **Step 7: Commit**

```bash
git add README.md docs/configuration.md
git commit -m "docs: document configuration and packaging"
```

---

## Self-Review Notes

- `dist/` and `node_modules/` are excluded by `.gitignore` from Task 1; the generated icons under `app/assets/icons/` and `build/` are committed deliberately (Task 10) because the build needs them.
- Every spec success criterion maps to a task: launch + configurable Outlook URL (16, 17), native notifications (14, 15), tray with best-effort badge (10–13), config precedence (3, 6, 7), AppImage/deb/rpm (1, 18), lint + unit tests (every task's final step, verified end-to-end in 18).
- Two files not named in the spec's project tree were added for testability: `app/browser/tools/notificationShim.js` (Task 15) and `app/ipc/register.js` (Task 17). Both exist so the logic they hold can be unit tested without a running Electron; `preload.js` and `index.js` stay as thin wiring.

