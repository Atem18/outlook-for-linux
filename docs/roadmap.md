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

- `mailto:` handler: register as the default mail client (desktop entry
  `MimeType=x-scheme-handler/mailto`, `setAsDefaultProtocolClient`), translate
  to `/mail/deeplink/compose?to=&subject=&body=&cc=&bcc=`, route through
  `second-instance`. Model: teams-for-linux `protocolLink.js` (`a85cef4c`)
- Calendar reminders while hidden (`backgroundThrottling` is already off;
  verify that reminders fire from the tray)
- Multiple accounts in isolated partitions (port `profilesManager`,
  `profileSwitcher`), e.g. work and personal side by side. When porting, also
  apply only the visible account's title-bar overlay (teams-for-linux
  `4a1605a1`) and keep one unread count per account
- Intune / Microsoft Identity Broker SSO, client certificates, WebAuthn (port
  `intune`, `clientCertificate`, `webauthn` with the wait-for-key fix
  `1266c3e1`). Intune needs `webRequest.onBeforeSendHeaders`, which app/unread
  already owns: share the listener
- Themes and custom CSS, global shortcuts, zoom persistence

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
