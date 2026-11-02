# CLAUDE.md

Outlook for Linux is an Electron wrapper around the Outlook web app, modelled on
teams-for-linux (local checkout at `../teams-for-linux`). When porting a module,
read the teams-for-linux version first and keep its idioms, but drop anything
Teams-specific.

## Environment

The host (SteamOS) has no Node. Run every npm/node/electron command inside the
`outlook-for-linux` distrobox (Arch, Node 26, shares $HOME and the display):

```bash
distrobox enter outlook-for-linux -- sh -c 'cd ~/Projects/outlook-for-linux && npm run test:unit'
```

The box needs `libnotify` for native notifications (installed with pacman).
npm 12 blocks dependency install scripts unless approved
(`npm install-scripts ls`). Electron downloads its binary on first run.

## Commands

- `npm start`: run in development
- `npm run lint`: ESLint, run before every commit
- `npm run test:unit`: `node --test 'tests/unit/*.test.js'`, run before every commit
- `npm run test:e2e`: Playwright launch test (needs a display or `xvfb-run`)
- `npm run dist:linux:deb` / `dist:linux:snap` / `dist:flatpak`: packages

## Layout

- `app/index.js`: entry point; keep it thin, put features in modules
- `app/config/`: option schema (`options.js`, plain data) and loader
- `app/appConfiguration/`: owns the frozen startup config and the settings store
- `app/mainAppWindow/`: the window; `linkPolicy.js` (URL policy), `startUrl.js`
  (remembered mailbox) and `titleBar.js` (integrated title bar) are pure and tested
- `app/browser/preload.js` + `tools/`: injected into Outlook. The preload runs
  sandboxed, so it is bundled by `scripts/buildPreload.js` into
  `preload.bundle.js` (git-ignored; built by `prestart`, `pretest:e2e`, packaging).
  Config reaches it over the synchronous `get-renderer-config` IPC, not
  `process.argv`. Page-world hooks (Notification, Badging API) go through
  `contextBridge.exposeInMainWorld` + `executeInMainWorld` before page scripts.
- `flatpak/`: local Flatpak manifest, desktop entry, AppStream metadata
- `docs/roadmap.md`: phases and what is next; update it when a phase item lands

## Rules

- `const`/`let` only, async/await, `#private` class fields, arrow callbacks
- Keep logic that can be pure in Electron-free modules so `node:test` covers it
- Never log PII: no email addresses, account ids, tokens, or full URLs (query
  strings carry tokens). Log schemes, origins, or counts instead.
- Main-frame https navigation must stay in-app: federated sign-in visits
  arbitrary identity-provider hosts. Email links arrive through window.open.
- Safe Links hosts are outlook.com subdomains but must open externally.
- `app/unread` borrows the bearer token Outlook sends to /owa/service.svc. It
  stays in main-process memory: never log it, never send it to a renderer,
  never send it to another origin. It owns the session's only
  `webRequest.onBeforeSendHeaders` listener (Electron allows one per session);
  anything else needing that hook must share it.
- To stop a dev instance use `pkill -f "[o]utlook-for-linux/node_modules/electron/dist/electron"`;
  without the brackets pkill matches and kills its own shell.
