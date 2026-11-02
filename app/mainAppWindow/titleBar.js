/**
 * Integrated title bar (client-side decorations), main-process half.
 * Ported from teams-for-linux (app/mainAppWindow/titleBar.js).
 *
 * With `titleBar: "integrated"` the main window hides the system title bar and
 * Electron draws native minimise / maximise / close buttons over the right end
 * of Outlook's own top bar. The top bar becomes the drag handle. The renderer
 * half (app/browser/tools/integratedTitleBar.js) measures that bar and reports
 * its height and colour, so the controls fit it.
 *
 * macOS is excluded: its traffic lights sit on the left, over Outlook's own
 * controls, and the system title bar there already matches the platform.
 */

// Outlook's top bar is 48px tall; the renderer corrects this once it has
// measured the real bar.
const DEFAULT_HEIGHT = 48;
const MIN_HEIGHT = 24;
const MAX_HEIGHT = 80;

// The overlay itself is transparent (alpha 00), so the buttons sit directly on
// whatever the page draws there (the Outlook top bar, the sign-in background,
// the loading screen) instead of on a block of guessed colour. The RGB part
// still matters: the hover and pressed backgrounds are derived from it, so a
// transparent black gives a dark hover square that hides dark glyphs on a
// light page. The renderer therefore reports the page colour, and the overlay
// uses it at zero alpha.
const TRANSPARENT_ALPHA = "00";
const DARK_PAGE = "#000000";
const LIGHT_PAGE = "#ffffff";
const LIGHT_GLYPHS = "#ffffff";
const DARK_GLYPHS = "#242424";

const HEX_COLOR_RE = /^#[0-9a-f]{6}$/i;

/**
 * Whether the integrated title bar applies to this config on this platform.
 * @param {object} config
 * @param {string} [platform]
 * @returns {boolean}
 */
function isIntegratedTitleBar(config, platform = process.platform) {
  return config?.titleBar === "integrated" && platform !== "darwin";
}

/**
 * BrowserWindow options that turn the integrated title bar on.
 * @param {object} config
 * @param {boolean} dark - whether the system prefers a dark theme
 * @param {string} [platform]
 * @returns {object} options to spread into the BrowserWindow constructor
 */
function titleBarWindowOptions(config, dark, platform = process.platform) {
  if (!isIntegratedTitleBar(config, platform)) {
    return {};
  }
  return {
    titleBarStyle: "hidden",
    titleBarOverlay: {
      color: (dark ? DARK_PAGE : LIGHT_PAGE) + TRANSPARENT_ALPHA,
      symbolColor: dark ? LIGHT_GLYPHS : DARK_GLYPHS,
      height: DEFAULT_HEIGHT,
    },
  };
}

/**
 * Validate an overlay update sent by the renderer.
 * @param {unknown} payload
 * @returns {{height: number, color: string, symbolColor: string, sample: boolean} | null}
 *   the overlay settings to apply (the page colour at zero alpha), and whether
 *   the renderer could not read a colour and the page must be sampled instead
 */
function sanitizeOverlayUpdate(payload) {
  if (!payload || typeof payload !== "object") {
    return null;
  }
  const { height, color, symbolColor, sample } = payload;
  if (
    !Number.isInteger(height) ||
    height < MIN_HEIGHT ||
    height > MAX_HEIGHT ||
    typeof color !== "string" ||
    !HEX_COLOR_RE.test(color) ||
    typeof symbolColor !== "string" ||
    !HEX_COLOR_RE.test(symbolColor)
  ) {
    return null;
  }
  return { height, color: color + TRANSPARENT_ALPHA, symbolColor, sample: sample === true };
}

// Relative luminance (WCAG) to pick light or dark control glyphs.
function isDark({ r, g, b }) {
  const lin = (c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b) < 0.4;
}

/**
 * Overlay colours from a sampled pixel of the page behind the controls.
 * Outlook paints its top bar with a theme image, which has no computed colour
 * the renderer could read, so the main process captures the page there.
 * @param {Buffer} bgra - one pixel in Chromium's native BGRA order
 * @returns {{color: string, symbolColor: string} | null}
 */
function overlayColorsFromPixel(bgra) {
  if (!bgra || bgra.length < 4) return null;
  const rgb = { r: bgra[2], g: bgra[1], b: bgra[0] };
  const hex = `#${[rgb.r, rgb.g, rgb.b].map((n) => n.toString(16).padStart(2, "0")).join("")}`;
  return {
    color: hex + TRANSPARENT_ALPHA,
    symbolColor: isDark(rgb) ? LIGHT_GLYPHS : DARK_GLYPHS,
  };
}

// The page area behind the window controls: the right end of the top bar.
// Wider than the controls themselves so the average is not one stray pixel.
const SAMPLE_WIDTH = 160;

function sampleRect(contentWidth, height) {
  const width = Math.min(SAMPLE_WIDTH, contentWidth);
  return { x: Math.max(0, contentWidth - width), y: 0, width, height };
}

module.exports = {
  DEFAULT_HEIGHT,
  isIntegratedTitleBar,
  titleBarWindowOptions,
  sanitizeOverlayUpdate,
  overlayColorsFromPixel,
  sampleRect,
};
