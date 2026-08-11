#!/usr/bin/env bash
# =============================================================================
#  outlook-for-linux — install / uninstall script
#  Usage:
#    ./deploy/install.sh            install
#    ./deploy/install.sh --uninstall
# =============================================================================
set -euo pipefail

# ── Colours ───────────────────────────────────────────────────────────────────
R='\033[0;31m'; G='\033[0;32m'; Y='\033[1;33m'; B='\033[0;34m'; N='\033[0m'

# ── Paths ─────────────────────────────────────────────────────────────────────
APP_NAME="outlook-for-linux"
INSTALL_LIB="${HOME}/.local/lib/${APP_NAME}"
INSTALL_BIN="${HOME}/.local/bin"
DESKTOP_DIR="${HOME}/.local/share/applications"
ICON_DIR="${HOME}/.local/share/icons/hicolor"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

# ── Uninstall ─────────────────────────────────────────────────────────────────
if [[ "${1:-}" == "--uninstall" ]]; then
  echo -e "${B}[Outlook]${N} Uninstalling ${APP_NAME}…"

  if systemctl --user is-active --quiet "${APP_NAME}" 2>/dev/null; then
    systemctl --user stop "${APP_NAME}"
  fi
  if systemctl --user is-enabled --quiet "${APP_NAME}" 2>/dev/null; then
    systemctl --user disable "${APP_NAME}"
  fi
  rm -f  "${SERVICE_FILE}"
  systemctl --user daemon-reload 2>/dev/null || true

  rm -f  "${INSTALL_BIN}/${APP_NAME}"
  rm -rf "${INSTALL_LIB}"
  rm -f  "${DESKTOP_DIR}/${APP_NAME}.desktop"
  rm -f  "${ICON_DIR}/256x256/apps/${APP_NAME}.png"
  rm -f  "${ICON_DIR}/128x128/apps/${APP_NAME}.png"
  rm -f  "${ICON_DIR}/48x48/apps/${APP_NAME}.png"
  rm -f  "${AUTOSTART_DIR}/${APP_NAME}.desktop"

  update-desktop-database "${DESKTOP_DIR}" 2>/dev/null || true
  gtk-update-icon-cache -f -t "${ICON_DIR}" 2>/dev/null || true

  echo -e "${G}✓  ${APP_NAME} uninstalled.${N}"
  exit 0
fi

# ── Banner ────────────────────────────────────────────────────────────────────
echo ""
echo -e "${B}╔══════════════════════════════════════╗${N}"
echo -e "${B}║   outlook-for-linux  installer        ║${N}"
echo -e "${B}╚══════════════════════════════════════╝${N}"
echo ""

# ── 1. Dependency checks ──────────────────────────────────────────────────────
echo -e "${B}[1/6]${N} Checking dependencies…"

if ! command -v node &>/dev/null; then
  echo -e "${R}[ERR]${N} node not found. Install Node.js 18+ from https://nodejs.org"
  exit 1
fi
if ! command -v npm &>/dev/null; then
  echo -e "${R}[ERR]${N} npm not found."
  exit 1
fi
if ! command -v rsync &>/dev/null; then
  echo -e "${R}[ERR]${N} rsync not found. Install it first:"
  echo -e "       Fedora/RHEL:  ${Y}sudo dnf install rsync${N}"
  echo -e "       Debian/Ubuntu: ${Y}sudo apt install rsync${N}"
  exit 1
fi

NODE_MAJ="$(node -e 'process.stdout.write(process.versions.node.split(".")[0])')"
if [[ "${NODE_MAJ}" -lt 18 ]]; then
  echo -e "${R}[ERR]${N} Node.js 18+ required (found $(node --version))."
  exit 1
fi
echo -e "     Node $(node --version) ✓"

# ── 2. Copy app to ~/.local/lib/ ──────────────────────────────────────────────
echo -e "${B}[2/6]${N} Installing app to ${INSTALL_LIB}…"
rm -rf "${INSTALL_LIB}"
mkdir -p "${INSTALL_LIB}"

rsync -a \
  --exclude='deploy/' \
  --exclude='.git/' \
  --exclude='node_modules/' \
  --exclude='*.log' \
  --exclude='.gitignore' \
  "${PROJECT_ROOT}/" "${INSTALL_LIB}/"

# ── 3. npm install inside the installed copy ──────────────────────────────────
echo -e "${B}[3/6]${N} Installing npm dependencies in ${INSTALL_LIB}…"
cd "${INSTALL_LIB}"
npm install --prefer-offline --no-audit --no-fund 2>&1 | \
  grep -E '(added|warn NPM|error)' | sed 's/^/     /' || true

# Verify the Electron binary is resolvable from the install location.
ELECTRON_BIN="$(node -e 'process.stdout.write(require("electron"))')"
if [[ ! -x "${ELECTRON_BIN}" ]]; then
  echo -e "${R}[ERR]${N} Could not find Electron binary at: ${ELECTRON_BIN}"
  echo -e "       Run 'npm install' manually inside ${INSTALL_LIB} and retry."
  exit 1
fi
echo -e "     Electron binary: ${ELECTRON_BIN} ✓"

# ── 4. Launcher script ────────────────────────────────────────────────────────
echo -e "${B}[4/6]${N} Creating launcher at ${INSTALL_BIN}/${APP_NAME}…"
mkdir -p "${INSTALL_BIN}"

# The launcher resolves the Electron binary at runtime so it stays correct
# after npm updates inside the install directory.
cat > "${INSTALL_BIN}/${APP_NAME}" << 'LAUNCHER'
#!/usr/bin/env bash
INSTALL_LIB="${HOME}/.local/lib/outlook-for-linux"
export ELECTRON_OZONE_PLATFORM_HINT="${ELECTRON_OZONE_PLATFORM_HINT:-auto}"
exec "$(cd "${INSTALL_LIB}" && node -e 'process.stdout.write(require("electron"))')" \
     "${INSTALL_LIB}/main.js" "$@"
LAUNCHER
chmod +x "${INSTALL_BIN}/${APP_NAME}"
echo -e "     ${INSTALL_BIN}/${APP_NAME} ✓"

# ── 5. Icons ──────────────────────────────────────────────────────────────────
echo -e "${B}[5/6]${N} Installing icons…"
ICON_SRC="${INSTALL_LIB}/assets/icon.png"

install_icon() {
  local size="$1"
  local px="${size%%x*}"
  local dest="${ICON_DIR}/${size}/apps"
  mkdir -p "${dest}"
  if command -v convert &>/dev/null; then
    convert "${ICON_SRC}" -resize "${px}x${px}" "${dest}/${APP_NAME}.png" 2>/dev/null \
      || cp "${ICON_SRC}" "${dest}/${APP_NAME}.png"
  else
    cp "${ICON_SRC}" "${dest}/${APP_NAME}.png"
  fi
  echo -e "     ${size} ✓"
}

install_icon "256x256"
install_icon "128x128"
install_icon "48x48"
gtk-update-icon-cache -f -t "${ICON_DIR}" 2>/dev/null || true

# ── 6. .desktop file ──────────────────────────────────────────────────────────
echo -e "${B}[6/6]${N} Registering .desktop entry…"
mkdir -p "${DESKTOP_DIR}"

cat > "${DESKTOP_DIR}/${APP_NAME}.desktop" << DESKTOP
[Desktop Entry]
Version=1.0
Type=Application
Name=Outlook for Linux
GenericName=Email Client
Comment=Microsoft Outlook Web Client
Exec=${INSTALL_BIN}/${APP_NAME} %U
Icon=${APP_NAME}
Terminal=false
StartupWMClass=${APP_NAME}
StartupNotify=true
Categories=Network;Email;
MimeType=x-scheme-handler/mailto;
Keywords=mail;email;outlook;microsoft;
DESKTOP

xdg-mime default "${APP_NAME}.desktop" x-scheme-handler/mailto 2>/dev/null || true
update-desktop-database "${DESKTOP_DIR}" 2>/dev/null || true
echo -e "     .desktop entry registered ✓"
echo -e "     mailto: handler → Outlook ✓"

# ── PATH warning ──────────────────────────────────────────────────────────────
if [[ ":${PATH}:" != *":${INSTALL_BIN}:"* ]]; then
  echo ""
  echo -e "${Y}[WARN]${N} ${INSTALL_BIN} is not in your PATH."
  echo -e "       Add to ~/.bashrc or ~/.zshrc:"
  echo -e "       ${Y}export PATH=\"\$HOME/.local/bin:\$PATH\"${N}"
fi

# ── Done ──────────────────────────────────────────────────────────────────────
echo ""
echo -e "${G}╔══════════════════════════════════════╗${N}"
echo -e "${G}║   Installation complete!             ║${N}"
echo -e "${G}╚══════════════════════════════════════╝${N}"
echo ""
echo -e "  Launch (terminal):   ${Y}${APP_NAME}${N}"
echo -e "  Launch (app menu):   search for ${Y}Outlook${N}"
echo -e "  Uninstall:           ${Y}./deploy/install.sh --uninstall${N}"
echo ""
echo -e "  ${B}To build distributable packages (AppImage / deb / rpm):${N}"
echo -e "  ${Y}make build${N}        — all formats"
echo -e "  ${Y}make build-rpm${N}    — rpm only  (requires rpmbuild)"
echo -e "  ${Y}make build${N} builds with electron-builder; output goes to dist/"
echo ""
echo -e "  ${B}KDE systray tip:${N} if the tray icon is missing, add the"
echo -e "  ${Y}AppIndicator${N} widget to your panel:"
echo -e "  Right-click panel → Add Widgets → search 'AppIndicator'"
echo ""
