# outlook-for-linux MVP — Design

**Date:** 2026-08-14
**Status:** Approved

## Purpose

Wrap the Outlook web client in an Electron desktop app for Linux, following the
same architectural conventions as [teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux),
which this project is modeled on. This spec covers the MVP only.

## Success criteria

- App launches, loads Outlook web (`https://outlook.office.com` by default,
  configurable), and is usable as a normal browser-wrapped desktop client.
- New-mail browser notifications from Outlook web are forwarded to native OS
  notifications.
- Tray icon is present; best-effort unread-count badge if the DOM exposes it.
- Config file support (`~/.config/outlook-for-linux/config.json`) with the
  same override precedence pattern as teams-for-linux (defaults → config file
  → CLI flags).
- Installable as AppImage/deb/rpm via electron-builder.
- `npm run lint` and unit tests pass.

## Scope decisions

- **Bootstrap strategy:** port and trim. The generalizable teams-for-linux
  modules (config loading, IPC allowlist, preload/tools split, window shell,
  logging/error-handling boilerplate) are copied in and adapted for Outlook;
  everything Teams-specific is left behind rather than ported and disabled.
- **Architecture conventions:** reused from day one — `AppConfiguration`
  class, `security/ipcValidator.js` allowlist, `browser/preload.js` +
  `browser/tools/` split — even though the MVP's actual feature surface is
  small. Low overhead now, avoids a restructuring pass later.
- **Outlook endpoint:** defaults to `https://outlook.office.com` (work/school).
  Configurable via `config.json`/CLI so users can point at
  `outlook.live.com` (personal) or a tenant-specific URL instead.
- **Packaging:** electron-builder configured for `AppImage`, `deb`, `rpm`
  targets from the start (mirrors teams-for-linux's `dist:linux` scripts).
- **License:** GPL-3.0-or-later, matching teams-for-linux.
- **Explicitly out of scope for MVP** (candidates for later sub-projects,
  each with its own brainstorm → spec → plan cycle): MQTT status publishing,
  Microsoft Graph API integration, multi-account profiles/partitions, custom
  CSS/backgrounds, screen sharing, WebAuthn/Intune SSO, spellcheck, global
  shortcuts, quick chat, auto-updater, E2E tests.

## Project structure

```
outlook-for-linux/
├── app/
│   ├── index.js                 # Main process entry — window, IPC, lifecycle
│   ├── appConfiguration/        # Config loading + validation (ported)
│   ├── mainAppWindow/           # BrowserWindow creation, loads Outlook URL
│   ├── browser/
│   │   ├── preload.js           # ipcRenderer bridge, module init
│   │   └── tools/                # DOM-facing scripts (tray badge sync, theme)
│   ├── security/
│   │   └── ipcValidator.js      # IPC channel allowlist (ported)
│   ├── startup/
│   │   └── commandLine.js       # CLI switches (ported, trimmed)
│   ├── notifications/           # OS notification bridging
│   ├── config/
│   │   └── defaults.js          # Default config values (incl. default URL)
│   └── utils/
│       └── logSanitizer.js      # PII-safe logging (ported)
├── build/                       # electron-builder assets (icons)
├── tests/
│   └── unit/
├── docs/
├── package.json
├── eslint.config.mjs
└── LICENSE (GPL-3.0-or-later)
```

## Components & data flow

- **`app/config/defaults.js`** — default config object, including
  `{ url: "https://outlook.office.com", trayIconEnabled: true,
  disableNotifications: false, ... }`.
- **`appConfiguration/`** — merges defaults, `~/.config/outlook-for-linux/config.json`,
  and CLI flags in that precedence order, then freezes the result. Treated as
  immutable for the rest of the process lifetime, same rule as
  teams-for-linux.
- **`app/index.js`** — reads `AppConfiguration` once at startup, creates the
  window via `mainAppWindow/`, registers the small set of IPC handlers this
  MVP needs, registers top-level `uncaughtException`/`unhandledRejection`
  handlers (including the network-error-pattern allowlist that avoids
  crashing on transient network blips, ported unchanged).
- **`mainAppWindow/`** — creates the `BrowserWindow`, points `loadURL` at
  `config.url`, wires the preload script.
- **`browser/preload.js` + `browser/tools/`** — runs inside the Outlook page.
  Intercepts the web `Notification` constructor (Outlook web already uses
  browser notifications for new mail) and forwards it to the main process
  over IPC to render as a native OS notification. A tray-badge tool reflects
  unread count into the tray icon on a best-effort basis if Outlook exposes
  it in the DOM/title — this is the most DOM-fragile piece, wrapped in
  try/catch with silent fallback since Outlook's DOM can change without
  notice.
- **`security/ipcValidator.js`** — allowlist covering only the channels this
  MVP actually uses (notification-forward, tray-update); grows as channels
  are added later, per teams-for-linux's own rule.

Primary data flow: **Outlook shows a notification → preload's Notification
shim catches it → IPC to main → main renders a native notification + updates
tray icon.**

## Error handling & logging

- `electron-log` for structured logging, PII rules carried over verbatim:
  never log the Outlook URL's query params, email addresses, or tokens.
- Top-level `uncaughtException`/`unhandledRejection` handlers ported as-is.
- All Outlook-DOM access in preload/tools wrapped in try/catch with silent
  fallback (log + no-op) rather than throwing.

## Testing & packaging

- **Unit tests** (`tests/unit/*.test.js`, Node's built-in `node --test`):
  config merging/precedence, IPC allowlist validation, log sanitizer — ported
  largely unchanged since the modules themselves are ported.
- **E2E:** deferred out of MVP.
- **Packaging:** electron-builder config targeting `AppImage`, `deb`, `rpm`.
- **Lint:** same ESLint config style ported (`const`/`let`, no `var`,
  async/await, `#private` fields).
