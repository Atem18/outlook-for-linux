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
| `defaultMailClient` | `false` | Make the app the default mail client at startup (`mailto:` links open a compose window). |
| `customCSSLocation` | `""` | CSS file injected into Outlook's pages. |
| `globalShortcuts` | `[]` | Shortcuts forwarded to Outlook while the app is not focused, e.g. `["Control+Shift+M"]`. On Wayland, start with `--enable-features=GlobalShortcutsPortal`. |
| `intune` | `false` | Single sign-on through the Microsoft Identity Broker on Intune-enrolled machines. |
| `intuneUser` | `""` | Broker account to use when it holds several. |
| `webauthn` | `false` | Hardware security keys (FIDO2) for sign-in, through `fido2-tools`. |
| `webauthnExtraOrigins` | `[]` | Extra `https` sign-in origins allowed to use security keys (your own identity provider). |
| `webauthnDebug` | `false` | Verbose security key logs. |
| `clientCertPath` | `""` | PKCS#12 client certificate to import at startup. |
| `clientCertPassword` | `""` | Its password. |
| `clientCertPinDialog` | `false` | PIN dialog for smartcard (PKCS#11) client certificates. |
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

### Default mail client

Every package declares the app as a `mailto:` handler, so it shows up in the
desktop's default-applications settings. Choosing it there makes "send email"
links from other apps open an Outlook compose window. Outside Flatpak and
Snap, `--defaultMailClient` (or `"defaultMailClient": true`) sets it at
startup instead, like `xdg-settings set default-url-scheme-handler mailto
outlook-for-linux.desktop`.

### Zoom

Ctrl+Plus, Ctrl+Minus, Ctrl+0 and Ctrl+wheel zoom Outlook; the zoom is kept
for the next launch.

## Building packages

```bash
npm run dist:linux:deb     # .deb (x64 + arm64)
npm run dist:linux:snap    # .snap (needs snapcraft)
npm run dist:flatpak       # local Flatpak install (needs flatpak + org.flatpak.Builder)
```

## License

GPL-3.0-or-later
