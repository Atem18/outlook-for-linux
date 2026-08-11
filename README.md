# outlook-for-linux

Outlook Web client , with:
- Systray icon with red unread-count badge (drawn PNG, no deps)
- Native KDE taskbar badge via `app.setBadgeCount()` — same mechanism as teams-for-linux
- Desktop notifications for new mail
- Persistent login (cookies survive restarts)
- Single-instance (second launch focuses the window)
- Hide to tray on close, mailto: handler

---

## Quick start (dev)

```bash
git clone https://github.com/kevinmesser/outlook-for-linux && cd outlook-for-linux
make install
make start          # auto-detects Wayland vs X11
make wayland        # force Wayland
```

## Install for your user

```bash
make deploy
```

The installer:
1. Runs `npm install`
2. Copies the app to `~/.local/lib/outlook-for-linux/`
3. Creates a launcher at `~/.local/bin/outlook-for-linux`
4. Registers icons in `~/.local/share/icons/hicolor/`
5. Installs a `.desktop` entry (required for `app.setBadgeCount()` on KDE)
6. Sets Outlook as your default `mailto:` handler

## Uninstall

```bash
make uninstall
# or
./deploy/install.sh --uninstall
```

---

## How unread count works

Mirrors the **teams-for-linux** architecture exactly:

1. `preload.js` — `MutationObserver` on `<title>`. Outlook Web sets the title to
   `(N) Inbox – you@domain.com – Outlook` when N > 0.
   Sends `tray-update` IPC with `{ count, flash }` on every change.

2. `main.js` — on receiving `tray-update`:
   - `app.setBadgeCount(n)` → **KDE native taskbar badge** via
     `com.canonical.Unity.LauncherEntry` D-Bus (same call teams-for-linux logs
     as `"Badge count set to 'N'"`)
   - `tray.setImage(...)` → regenerates the systray PNG with a red badge overlay

3. Title-poll fallback every 5 s via `webContents.getTitle()` — catches anything
   the MutationObserver might miss after SPA navigations.

---

## KDE Plasma 6 / Wayland notes

### Systray badge
- `tray.setImage()` uses `nativeImage.createFromBuffer()` with a real PNG — the
  only method that works reliably on KDE Linux (SVG data URLs do not work)
- You need the **AppIndicator** (KStatusNotifierItem) widget in your panel:
  Right-click panel → Add Widgets → search **AppIndicator**

### Taskbar badge
- `app.setBadgeCount()` requires the `.desktop` file to be installed and its
  `StartupWMClass` to match the app name. `make deploy` handles this.

### Window decorations
- `WaylandWindowDecorations` is enabled so KDE renders native server-side borders

---

## Files

| File | Purpose |
|---|---|
| `main.js` | Main process — window, tray, IPC, Wayland flags, `setBadgeCount` |
| `preload.js` | Renderer bridge — MutationTitle → `tray-update` IPC |
| `icon-generator.js` | Pure-Node PNG renderer (zlib only, zero deps) |
| `assets/icon.png` | 256×256 app icon (generated at build time) |
| `deploy/install.sh` | Installer / uninstaller |
| `Makefile` | Convenience targets |
| `outlook-for-linux.desktop` | Reference `.desktop` file |
