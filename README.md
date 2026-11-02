# Outlook for Linux

Unofficial Microsoft Outlook client for Linux. It wraps the Outlook web app
in an Electron window and adds desktop integration, following the approach of
[teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux).

Not affiliated with or endorsed by Microsoft.

## Status

Early development (Phase 1). See [docs/roadmap.md](docs/roadmap.md).

## Running from source

Requires Node.js 22 or later.

```bash
npm install
npm start
```

## Configuration

Options come from, in increasing precedence:

1. `/etc/outlook-for-linux/config.json` (system-wide)
2. `~/.config/outlook-for-linux/config.json` (user)
3. Environment variables prefixed `OUTLOOK_FOR_LINUX_` (e.g. `OUTLOOK_FOR_LINUX_MINIMIZED=true`)
4. Command-line flags (e.g. `--minimized`)

| Option | Default | Description |
|---|---|---|
| `url` | `https://outlook.cloud.microsoft/mail/` | Outlook URL to load. |
| `partition` | `persist:outlook` | Session partition. Use different values to run separate accounts side by side. |
| `chromeUserAgent` | Chrome UA | User agent sent to Outlook. |
| `minimized` | `false` | Start minimized. |
| `titleBar` | `system` | `integrated` hides the desktop title bar and draws the window buttons over the Outlook top bar (client-side decorations). |
| `trayIcon` | `true` | Tray icon with the unread count. |
| `closeToTray` | `true` | Closing the window hides it to the tray, so notifications and reminders keep arriving. |
| `notifications` | `true` | Show Outlook's notifications as native desktop notifications. |
| `menubarAutoHide` | `true` | Hide the menu bar until Alt is pressed. |
| `spellCheckerLanguages` | `[]` | Spell checker languages; empty uses the system locale. |

### Accounts

Work, school and personal (outlook.com, hotmail, live) accounts all use the
same URL: Microsoft's sign-in page detects the account type and sends you to
the right mailbox. The app then remembers which mailbox you use and opens it
directly on the next launch.

To keep two accounts signed in at once, run a second instance with its own
data directory, as teams-for-linux does:

```bash
outlook-for-linux --user-data-dir="$HOME/.config/outlook-for-linux-personal"
```

## Building packages

```bash
npm run dist:linux:deb     # .deb (x64 + arm64)
npm run dist:linux:snap    # .snap (needs snapcraft)
npm run dist:flatpak       # local Flatpak install (needs flatpak + org.flatpak.Builder)
```

## License

GPL-3.0-or-later
