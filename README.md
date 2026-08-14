# Outlook for Linux

Unofficial desktop client for the Microsoft Outlook web app, built with
Electron. It wraps outlook.office.com (or outlook.live.com, or your tenant URL)
in a standalone window with a tray icon, an unread-count badge and native
desktop notifications.

Not affiliated with, endorsed by, or supported by Microsoft. "Outlook" and
"Microsoft" are trademarks of Microsoft Corporation; this project only loads
their web app in an Electron window.

Modelled on [teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux).

## What it does

- Loads the Outlook web app in its own window, with your sign-in session
  persisted between runs.
- Shows a tray icon with an unread-count badge, derived from the page title.
- Renders Outlook's web notifications as native desktop notifications, so they
  survive the window being backgrounded and follow your desktop's own
  notification settings. Clicking one raises the window.
- Closes to the tray by default; quit from the tray menu.
- Sends third-party links to your normal browser, while keeping Outlook and
  Microsoft sign-in pages in the app.

## Install

Download a package from the
[latest release](https://github.com/sandesh-suresh/outlook-for-linux/releases/latest),
or build one yourself (see [Building](#building)), then:

**AppImage** — no install needed:

```bash
chmod +x outlook-for-linux-*.AppImage
./outlook-for-linux-*.AppImage
```

**deb** (Debian, Ubuntu):

```bash
sudo apt install ./outlook-for-linux_*_amd64.deb
```

**rpm** (Fedora, RHEL, openSUSE):

```bash
sudo dnf install ./outlook-for-linux-*.x86_64.rpm
```

## Configuration

All options, defaults and precedence rules are documented in
[docs/configuration.md](docs/configuration.md).

The short version: create `~/.config/outlook-for-linux/config.json`:

```json
{
  "url": "https://outlook.office.com",
  "trayIconEnabled": true
}
```

Any option can also be passed as a CLI flag, e.g.
`--url=https://outlook.live.com`. CLI flags take precedence over the config
file, which takes precedence over the defaults. Configuration is read once at
startup, so changes need a restart.

## Development

```bash
npm install
npm start          # run the app
npm run lint       # ESLint
npm run test:unit  # unit tests
```

## Building

```bash
npm run pack        # unpacked build in dist/linux-unpacked/
npm run dist:linux  # AppImage, deb and rpm in dist/
```

The rpm target needs `rpmbuild` (`sudo dnf install rpm-build`).

## Continuous integration

Every pull request runs ESLint, the unit tests and `npm audit`, then builds the
AppImage, deb and rpm and attaches them to the run as a downloadable artifact,
so a change can be installed and tried before it is merged. CodeQL and
OSV-Scanner check the code and the dependency tree; Dependabot opens weekly
dependency PRs.

## Releases

Versioning is driven by commit messages, so releases need no version editing by
hand. Commit in [conventional commit](https://www.conventionalcommits.org/)
form — `feat:`, `fix:`, `docs:`, `chore:` — and
[release-please](https://github.com/googleapis/release-please) keeps a
`chore: release X.Y.Z` pull request open on `main` with the computed version and
the assembled `CHANGELOG.md`. Nothing is released until that PR is merged;
merging it bumps `package.json`, creates the `vX.Y.Z` tag and the GitHub
release, and the build workflow attaches the packages to it.

`feat:` bumps the minor version and `fix:` the patch version. A breaking change
(`feat!:`, or a `BREAKING CHANGE:` footer) bumps the minor version too while the
project is pre-1.0.

## Not in this release

Deliberately out of scope for the MVP, each a candidate for its own later
change: MQTT status publishing, Microsoft Graph API integration, multi-account
profiles, custom CSS and backgrounds, screen sharing, WebAuthn/Intune SSO,
spellcheck, global shortcuts, quick chat, an auto-updater, and end-to-end
tests.

## Licence

GPL-3.0-or-later. See [LICENSE](LICENSE).
