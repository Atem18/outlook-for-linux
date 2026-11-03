# Roadmap

Outlook for Linux follows the architecture of
[teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux): an
Electron main process with one module per concern under `app/`, wrapping the
Outlook web app. Generic modules are ported from teams-for-linux as each phase
needs them; Teams-specific ones (calls, screen sharing, camera, MQTT, Quick
Chat, stickers) are not. When porting, read the teams-for-linux module first
(local checkout at `../teams-for-linux`) and keep its idioms.

## Decisions

- New repository, not a fork; plain JavaScript (no TypeScript)
- Work, school and personal accounts from day one, through one account-neutral
  URL, as teams-for-linux does (no `account` option)
- Packaged as Flatpak (`io.github.Atem18.OutlookForLinux`), deb and snap
- Main window sandboxed with contextIsolation; the preload is bundled with
  esbuild because a sandboxed preload can only require `electron`
- Integrated title bar (CSD) available but off by default (`titleBar: "system"`),
  as in teams-for-linux
- Development runs in the `outlook-for-linux` distrobox (see CLAUDE.md)

## Phase 0: scaffold (done)

- Repository, lint, `node:test` unit tests, Playwright launch test, CI
- Config system: system + user `config.json`, env vars, CLI flags (yargs)
- One account-neutral URL (`outlook.cloud.microsoft`) for work and personal
  accounts; separate instances via `--user-data-dir`
- Main window: persisted window state, Chrome user agent, single instance lock
- Link policy: Outlook pop-outs stay in the app, email links and Safe Links open
  in the browser, federated sign-in stays in the app
- Remembered mailbox: reopens outlook.live.com or the work mailbox directly
  (remembered only once the page stays on the mailbox)
- Sign-out: Microsoft's signed-out dead ends (`/logout`, `/logoutsession`, the
  marketing site, msn.com after a personal sign-out) go back to sign-in and
  forget the mailbox (backport of teams-for-linux `188dd7fc`)
- Integrated title bar (`titleBar: "integrated"`), ported from the
  teams-for-linux CSD work (`765e758c`, `4a1605a1`); glyph colour sampled from
  the page when Outlook paints its bar with a theme image
- Packaging config for deb, snap and a local Flatpak

## Phase 1: MVP (done)

- Tray icon with close-to-tray (`trayIcon`, `closeToTray`); Quit in its menu
- Unread badge (app/unread): a read-only GetFolder on the Inbox through
  Outlook's own API (/owa/service.svc), every 30 s, with the bearer token the
  Outlook web app sends on its own calls, borrowed like teams-for-linux's
  graphApi borrows the Teams Graph token (kept in main-process memory, never
  logged, only sent back to its origin). Fallback: the Inbox count in the
  folder pane, which Outlook does not render in narrow windows. The page draws
  the badge; the tray composites it onto the icon. Outlook's Badging API calls
  are not the total (it clears the badge on every load)
- Native notifications: the page's `Notification` is replaced through
  `contextBridge.executeInMainWorld`; clicks bring the window back and reach
  Outlook's handler. Needs libnotify at run time (the deb depends on it; the
  Flatpak runtime ships it)
- Reload with backoff when Outlook failed to load for lack of network, and on
  resume from suspend
- Download progress in the taskbar and a "download complete" notification

### Phase 1 follow-ups

- Verify on a personal account (outlook.live.com): unread count through the
  API, folder-pane fallback selector, sign-out flow
- Verify close-to-tray, Outlook-sent notifications (Outlook settings →
  Notifications → Desktop) and downloads by hand
- KDE/GNOME taskbar badge: Unity LauncherEntry over D-Bus, as teams-for-linux
  does (`app.setBadgeCount` only works on Unity)
- Notification sounds
- Unread count lags up to 30 s; consider refreshing right after Outlook's own
  `GetFolderChangeDigest`/notification-channel activity

## Phase 2: desktop integration

- `mailto:` handler (done): every package declares `x-scheme-handler/mailto`;
  a `mailto:` link on the command line (first launch or `second-instance`)
  opens Outlook's `/mail/deeplink/compose` in a pop-out on the current or
  remembered mailbox, so the main window keeps its state. Parsing follows
  RFC 6068 (`+` stays literal). Claiming the default is opt-in
  (`defaultMailClient`, skipped inside Flatpak/Snap). Model: teams-for-linux
  `protocolLink.js` (`a85cef4c`)
  - Cc/Bcc: the deeplink ignores `cc`/`bcc` (verified on a work mailbox).
    `?mailtouri=<raw link>` (what Outlook registers as its browser handler)
    fills them, but then shows spaces in subject and body as "+", and other
    parameters are ignored next to it. Untested: `Cc`/`Bcc` capitalised,
    literal spaces inside `mailtouri`
  - Integrated title bar: the compose window gets a 32px title strip above
    the ribbon, as Outlook for Windows gives its pop-outs, in the main
    window's header colour (sampled from its top rows) with the window
    controls in it (`app/browser/tools/popOutTitleBar.js`). Outlook's own
    pop-outs (about:blank windows Outlook writes into) still use the system
    title bar. When a mailto: link starts the app, the strip may take the
    loading screen's colour until the main window reports its bar again
  - Follow-up: verify by hand on signed-in work and personal profiles (the
    compose page, what Outlook shows after sending, links from Firefox and
    `xdg-email`; attachments wait for "send to mail recipient" in Phase 3)
- Calendar reminders while hidden (`backgroundThrottling` is already off;
  verify that reminders fire from the tray): still to verify by hand
- Multiple accounts in isolated partitions (port `profilesManager`,
  `profileSwitcher`), e.g. work and personal side by side. When porting, also
  apply only the visible account's title-bar overlay (teams-for-linux
  `4a1605a1`) and keep one unread count per account. Until then, a second
  instance with `--user-data-dir` (README) runs a second account
- Intune SSO (done, `intune`, `intuneUser`): the Microsoft Identity Broker's
  PRT SSO cookie on login.microsoftonline.com requests, over D-Bus. The
  session's one `onBeforeSendHeaders` listener is now shared through
  `app/requestHeaders` (unread and Intune). Not logged from teams-for-linux:
  the broker's raw account list. Untested here (no Intune-enrolled machine)
- Client certificates (done): `clientCertPath` imports a PKCS#12 file;
  `clientCertPinDialog` shows the smartcard PIN dialog (shared hardened
  prompt in `app/_shared/securePrompt.js`). Untested here (no card)
- WebAuthn / FIDO2 security keys (done, `webauthn`, `webauthnExtraOrigins`):
  fido2-tools backend, PIN and touch prompts and the wait-for-key fix
  (`1266c3e1`) ported as is. The window isolates its preload, so the
  `navigator.credentials` patch runs in the page world through
  `executeInMainWorld`, only on the allowed sign-in origins; the relay for
  sign-in iframes stays in the isolated world. deb recommends fido2-tools;
  the Flatpak would need `--device=all` and fido2-tools in the bundle.
  Untested with a real key
- Zoom persistence (done): View menu and Ctrl+wheel zoom, saved for the next
  launch; Chromium applies it to every window of the same origin
- Global shortcuts (done, `globalShortcuts`), ported as is
- Custom CSS (done, `customCSSLocation`). Teams' packaged themes restyle
  Teams and are not ported; Outlook's own themes and dark mode (following the
  desktop's light/dark preference) cover theming

## Phase 3: extras

- Microsoft Graph through the page's session: a reminder backstop
- Open `.eml` and `.ics` files; "send to mail recipient" with attachments via a
  draft created through Outlook's API
- Auto-update for deb (electron-updater)
- Flathub submission with vendored npm sources (flatpak-node-generator); Snap
  Store
- Copy a chosen app icon into the user icon theme (teams-for-linux `f67da455`)

## Could go back to teams-for-linux

- Page sampling behind the window controls: on the Microsoft sign-in pages,
  whose background is an image, the integrated title bar falls back to the
  system scheme and draws white glyphs on a light page in dark mode
- Sandboxed, bundled preload: Chromium may start the page in a process created
  for the site's service worker, which is sandboxed whatever the window asks
  for, so a preload that requires local files can fail there
