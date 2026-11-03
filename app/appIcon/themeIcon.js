const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

// Copies a chosen app icon into the per-user icon theme. Ported from
// teams-for-linux app/menus/themeIcon.js (f67da455). Its window class is
// fixed (WINDOW_CLASS) until multiple instances get a --class option.
//
// The desktop resolves taskbar, launcher and app menu icons by name from the
// icon theme, never from the running window, so a custom icon only shows
// there once a copy sits in the per-user hicolor theme under that name.
//
// Electron-free: the caller passes a nativeImage (anything with getSize and
// resize().toPNG()), so node:test covers it.
const SIZES = [16, 22, 24, 32, 48, 64, 96, 128, 256, 512];
const DEFAULT_NAME = "outlook-for-linux";
// The window's class (X11 WM_CLASS, Wayland app_id), from desktopName in
// package.json. The shell matches a launcher to the window by it.
const WINDOW_CLASS = "outlook-for-linux";
const NAME_PATTERN = /^[\w.-]+$/;

let bus = null;

function dataHome() {
  return process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local", "share");
}

// Inside Flatpak, XDG_DATA_HOME is the sandbox's own (~/.var/app/<id>/data);
// the host's icon theme is mounted at its usual path by the manifest's
// --filesystem=xdg-data/icons:create.
function iconsDir() {
  const hostDataHome = process.env.FLATPAK_ID
    ? process.env.HOST_XDG_DATA_HOME || path.join(os.homedir(), ".local", "share")
    : dataHome();
  return path.join(hostDataHome, "icons");
}

function hicolorDir() {
  return path.join(iconsDir(), "hicolor");
}

function manifestFile() {
  return path.join(dataHome(), DEFAULT_NAME, "theme-icon.json");
}

// A Snap cannot write the host's icon theme. A Flatpak can when its manifest
// grants the icons directory (ours does; a build without it cannot).
function supported() {
  if (process.platform !== "linux" || process.env.SNAP) return false;
  if (!process.env.FLATPAK_ID) return true;
  try {
    fs.accessSync(iconsDir(), fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

function readDesktopFiles() {
  const dirs = [dataHome(), ...(process.env.XDG_DATA_DIRS || "/usr/local/share:/usr/share").split(":")];
  const files = [];
  for (const dir of dirs) {
    const appsDir = path.join(dir, "applications");
    let entries;
    try {
      entries = fs.readdirSync(appsDir);
    } catch {
      continue;
    }
    for (const entry of entries) {
      if (!entry.endsWith(".desktop")) continue;
      try {
        files.push({ name: entry, text: fs.readFileSync(path.join(appsDir, entry), "utf8") });
      } catch {
        // unreadable entries are skipped
      }
    }
  }
  return files;
}

// The icon name of the desktop file the shell matches this window to: the
// launcher whose StartupWMClass is the window class (as the packaged entry
// and hand-made launchers set it), otherwise the packaged entry by file name,
// otherwise whichever launches this binary (AppImage integrations install one
// under their own icon name). A value that is not a plain name (a path, "..")
// falls back, so it can never write outside the theme.
function iconName() {
  // A Flatpak's exported launcher is named after its id and uses it as its
  // icon; the user theme takes precedence over the exported icons.
  const flatpakId = process.env.FLATPAK_ID;
  if (flatpakId) return NAME_PATTERN.test(flatpakId) ? flatpakId : DEFAULT_NAME;
  const files = readDesktopFiles();
  const exe = process.env.APPIMAGE || process.execPath;
  const match =
    files.find((f) => f.text.split("\n").some((l) => l.trim() === `StartupWMClass=${WINDOW_CLASS}`)) ||
    files.find((f) => f.name === `${DEFAULT_NAME}.desktop`) ||
    files.find((f) => f.text.includes(exe));
  const icon = match && /^Icon=(.+)$/m.exec(match.text)?.[1].trim();
  return icon && NAME_PATTERN.test(icon) ? icon : DEFAULT_NAME;
}

function iconFile(size, name) {
  return path.join(hicolorDir(), `${size}x${size}`, "apps", `${name}.png`);
}

// One manifest per user, keyed by icon name, recording the files this app
// wrote so removal never touches anything else.
function readManifest() {
  try {
    const data = JSON.parse(fs.readFileSync(manifestFile(), "utf8"));
    return Array.isArray(data) ? {} : data;
  } catch {
    return {};
  }
}

function writeManifest(manifest) {
  if (Object.keys(manifest).length === 0) {
    fs.rmSync(manifestFile(), { force: true });
    return;
  }
  fs.mkdirSync(path.dirname(manifestFile()), { recursive: true });
  fs.writeFileSync(manifestFile(), JSON.stringify(manifest));
}

// fs errors carry the full path, which includes the user name: log the code.
function warn(message, error) {
  console.warn(`[APP_ICON] ${message}`, { code: error.code ?? error.name });
}

function refresh() {
  // GTK desktops watch the theme directory's mtime; KDE listens for this signal.
  const now = new Date();
  try {
    fs.utimesSync(hicolorDir(), now, now);
  } catch (error) {
    warn("Could not touch the icon theme directory", error);
  }
  try {
    if (!bus) {
      // A failed connect arrives as an async error event, which this catch
      // never sees; without the listener it would take the app down.
      bus = require("@homebridge/dbus-native").sessionBus();
      bus.connection.on("error", (error) => {
        warn("Session bus connection failed", error);
        bus = null;
      });
    }
    bus.sendSignal("/KIconLoader", "org.kde.KIconLoader", "iconChanged", "i", [0]);
  } catch (error) {
    warn("Could not send the icon change signal", error);
  }
}

function installed() {
  return supported() && Boolean(readManifest()[iconName()]);
}

// Every size goes to a temp name first and is renamed over the real file only
// once all of them are written, so a failed replacement leaves the previous
// icon set whole. A size that already holds a file this app did not write is
// skipped, so hand-placed icons are never overwritten or later removed.
function install(image) {
  if (!supported()) return false;
  const name = iconName();
  const owned = new Set(readManifest()[name]);
  const { width, height } = image.getSize();
  const pending = [];
  try {
    for (const size of SIZES) {
      const file = iconFile(size, name);
      if (!owned.has(file) && fs.existsSync(file)) continue;
      const temp = `${file}.tmp-${process.pid}`;
      fs.mkdirSync(path.dirname(file), { recursive: true });
      const resized = image.resize(width >= height ? { width: size } : { height: size });
      fs.writeFileSync(temp, resized.toPNG());
      pending.push({ temp, file });
    }
    if (pending.length === 0) {
      console.warn("[APP_ICON] Every size already holds an icon this app did not write, leaving them alone");
      return false;
    }
    for (const { temp, file } of pending) fs.renameSync(temp, file);
    writeManifest({ ...readManifest(), [name]: pending.map((p) => p.file) });
  } catch (error) {
    warn("Could not write the icon theme, leaving it untouched", error);
    for (const { temp } of pending) fs.rmSync(temp, { force: true });
    return false;
  }
  refresh();
  return true;
}

// Only files this app wrote under this name are removed.
function remove() {
  if (!supported()) return false;
  const manifest = readManifest();
  const name = iconName();
  const files = manifest[name];
  if (!files) return false;
  try {
    for (const file of files) fs.rmSync(file, { force: true });
    delete manifest[name];
    writeManifest(manifest);
  } catch (error) {
    warn("Could not remove the icon theme files", error);
    return false;
  }
  refresh();
  return true;
}

module.exports = { install, installed, remove, iconName, supported };
