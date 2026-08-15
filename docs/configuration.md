# Configuration

Configuration is read once at startup from
`~/.config/outlook-for-linux/config.json`. Command-line flags override the
file, and the file overrides the defaults. Changes require a restart.

The file is optional — without it, every option below takes its default.
A malformed file is reported in the log and ignored in favour of the defaults,
so a typo never stops the app from starting.

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

## Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `url` | string | `https://outlook.office.com` | Outlook web URL. Use https://outlook.live.com for personal accounts, or a tenant-specific URL. |
| `appTitle` | string | `Microsoft Outlook` | Window and tray tooltip title. Outlook replaces the window title with its own once the page loads. |
| `appIcon` | string | `""` (empty) | Absolute path to a custom tray icon. Empty means use the bundled icon. |
| `appIconType` | string | `default` | Which bundled tray icon variant to use. One of `default`, `light`, `dark` — the monochrome `light` and `dark` variants suit panels that expect a single-colour icon. |
| `trayIconEnabled` | boolean | `true` | Enable the tray icon. |
| `closeToTray` | boolean | `true` | Hide to tray instead of quitting when the window is closed. Ignored when `trayIconEnabled` is `false`, since hiding with no tray would leave no way back to the window. |
| `disableNotifications` | boolean | `false` | Disable all native notifications. |
| `defaultNotificationUrgency` | string | `normal` | Urgency passed to the native notification server. One of `low`, `normal`, `critical`. |
| `useUnreadListLogic` | boolean | `true` | Watch the mail list with a MutationObserver to derive the unread count for the tray badge. |
| `disableBadgeCount` | boolean | `false` | Do not draw the unread count badge on the tray icon. Also suppresses the taskbar/dock badge. |
| `disableNotificationWindowFlash` | boolean | `false` | Do not flash the window frame when the unread count increases. |
| `disableGpu` | boolean | `false` | Disable GPU hardware acceleration. Try this first if the window renders blank or the app crashes on startup. |
| `logConfig` | object | `{"transports":{"console":{"level":"info"},"file":{"level":false}}}` | electron-log configuration. Set to the string "console" for plain console logging, or false to disable logging. The object is passed through to electron-log as-is; `"console"` leaves the global console untouched, and `false` silences logging entirely. |
| `watchConfigFile` | boolean | `false` | Log a warning when the config file changes. Config is read once at startup, so a restart is required to apply changes. |
| `webDebug` | boolean | `false` | Open DevTools on start. |

## Precedence

1. Command-line flags — highest.
2. `~/.config/outlook-for-linux/config.json`.
3. The defaults above — lowest.

The resolved configuration is frozen at startup and never re-read, so the
running app cannot drift from what the log reported at launch.

## Logging

Logs go to the console at `info` level by default. To raise the level for a
single run:

```bash
outlook-for-linux --logConfig.transports.console.level=debug
```

Log output is filtered for personal information before it is written, so
mailbox addresses, tokens and URL query parameters do not end up in a log you
might attach to a bug report.
