# Outlook for Linux

Unofficial Microsoft Outlook client for Linux. It wraps the Outlook web app
in an Electron window and adds desktop integration, following the approach of
[teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux).

Not affiliated with or endorsed by Microsoft.

## Status

Early development (Phase 2). See [docs/roadmap.md](docs/roadmap.md).

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
| `titleBar` | `integrated` | `integrated` hides the desktop title bar and draws the window buttons over the Outlook top bar (client-side decorations); `system` keeps the desktop's title bar. |
| `appIcon` | `""` | PNG used instead of the bundled icon (windows, tray, notifications). See [App icon](#app-icon). |
| `trayIcon` | `true` | Tray icon with the unread count (the taskbar/dock badge shows it either way). |
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

### App icon

App Icon → Choose App Icon…, in the View menu and the tray menu, swaps the icon of the windows, the tray
and notifications for a PNG of your own, at once and for the next launches;
Reset to default goes back. The menu stores its choice as the `appIcon`
option in `config.json`, like teams-for-linux, so it can be set by hand there
too or for one launch on the command line (`--appIcon=/path/to/icon.png`),
handy for telling several instances apart. The menu leaves a `config.json`
that is not valid JSON alone and says so.

The menu keeps a copy of the chosen PNG (`app-icon.png`, next to
`config.json`), so moving or deleting the original later is fine. In the Flatpak, a path given to
`appIcon` must be one the sandbox can read, such as `~/Downloads`.

The taskbar, launcher and app menu take their icon from the icon theme, not
from the running app. The app offers to copy the chosen icon into
`~/.local/share/icons` (the "Also use it for the taskbar and launcher"
checkbox), under the launcher's icon name; the Flatpak may write only that
folder of `~/.local/share`. It shows from the next start, on some desktops
after logging in again. Reset removes the copied files, and only those. Not
available in the Snap.

### Zoom

Ctrl+Plus, Ctrl+Minus, Ctrl+0 and Ctrl+wheel zoom Outlook; the zoom is kept
for the next launch.

## Building packages

```bash
npm run dist:linux:deb     # .deb (x64 + arm64)
npm run dist:linux:snap    # .snap (needs snapcraft)
npm run dist:flatpak       # local Flatpak install (needs flatpak + org.flatpak.Builder)
```

Packages land in `dist/`. The `.deb` names end in `_amd64.deb` (Intel/AMD)
or `_arm64.deb` (ARM); build for the machine's architecture with
`npx electron-builder --x64 -l deb --publish never` (or `--arm64`) to skip
the other one.

## Installing on Ubuntu or Debian

There is no download yet, so build the `.deb` first, on any Linux machine
with Node.js 22 or later (it does not have to be the one you install on):

```bash
git clone https://github.com/Atem18/outlook-for-linux.git
cd outlook-for-linux
npm install
npm run dist:linux:deb
```

If packaging fails with `libcrypt.so.1: cannot open shared object file`,
install the libcrypt compatibility package (`libxcrypt-compat` on Arch and
Fedora; Ubuntu and Debian already have it) and run the last command again.

Copy `dist/outlook-for-linux_<version>_amd64.deb` to the target machine and
install it with apt, which also installs the libraries it needs:

```bash
sudo apt install ./outlook-for-linux_2.1.0_amd64.deb
```

Outlook for Linux is then in the applications menu, or run
`outlook-for-linux` from a terminal. The package's dependencies cover
Ubuntu 22.04 and later and Debian 12 and later. To make it the mail client for `mailto:` links, pick it
in Settings → Default Applications → Mail, or run:

```bash
xdg-mime default outlook-for-linux.desktop x-scheme-handler/mailto
```

The package does not update itself: to upgrade, build the new version and
install it the same way. To uninstall:

```bash
sudo apt remove outlook-for-linux
```

Settings and the signed-in session stay in `~/.config/outlook-for-linux`;
delete that folder too to remove everything.

## License

GPL-3.0-or-later
