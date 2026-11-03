/**
 * Integrated title bar (client-side decorations), renderer half.
 * Ported from teams-for-linux (app/browser/tools/integratedTitleBar.js).
 *
 * With `titleBar: "integrated"` the system title bar is hidden and Electron
 * draws the window controls over the right end of Outlook's top bar. This
 * module:
 *
 * - makes the Outlook top bar the window's drag region, while its buttons,
 *   search box and links stay clickable,
 * - reports the bar's height and colour, so the window controls (drawn on a
 *   transparent background) get matching size, glyph colour and hover tint,
 * - reserves room on the right when Outlook's own buttons would sit under the
 *   controls, unless Outlook already leaves that room itself through Window
 *   Controls Overlay,
 * - adds a thin drag strip on pages without an Outlook top bar (the Microsoft
 *   sign-in pages, the loading screen), so the window can still be moved
 *   there, hides the page scrollbar there so it does not run up under the
 *   controls, and picks the glyph colour from the page background.
 *
 * Outlook's DOM changes without notice, so every lookup is defensive and the
 * whole module degrades to "controls over the page" if nothing matches.
 */

const CHANNEL = "title-bar-overlay-update";
const STYLE_ID = "ofl-integrated-title-bar";
const DRAG_STRIP_ID = "ofl-title-bar-drag-strip";
const NO_TOP_BAR_CLASS = "ofl-no-top-bar";
const DRAG_CLASS = "ofl-title-bar-drag";
const RESERVE_CLASS = "ofl-wco-reserve";
const CHECK_DELAY_MS = 500;
const RESAMPLE_INTERVAL_MS = 10_000;
const DEFAULT_HEIGHT = 48;
const MIN_HEIGHT = 24;
const MAX_HEIGHT = 80;
const LIGHT_GLYPHS = "#ffffff";
const DARK_GLYPHS = "#242424";
const LIGHT_PAGE = { r: 255, g: 255, b: 255 };
const DARK_PAGE = { r: 0, g: 0, b: 0 };
const DARK_SCHEME = "(prefers-color-scheme: dark)";

// Controls known to live in Outlook's top bar, most specific first. The bar
// itself is found by walking up from whichever matches (see findTopBar).
const ANCHOR_SELECTORS = [
  "#O365_NavHeader",
  '[role="banner"]',
  "#topSearchInput",
  '[role="search"]',
  "#O365_MainLink_NavMenu",
];

const INTERACTIVE_SELECTOR =
  'button, a, input, select, textarea, [role="button"], [role="search"], [role="menuitem"], [role="tab"], [tabindex]';

const STYLE = `
  .${DRAG_CLASS} {
    -webkit-app-region: drag;
  }
  .${DRAG_CLASS} :is(${INTERACTIVE_SELECTOR}) {
    -webkit-app-region: no-drag;
  }
  .${RESERVE_CLASS} {
    box-sizing: border-box !important;
    padding-right: calc(100vw - env(titlebar-area-x, 0px) - env(titlebar-area-width, 100vw)) !important;
  }
  #${DRAG_STRIP_ID} {
    position: fixed;
    top: 0;
    left: env(titlebar-area-x, 0px);
    width: env(titlebar-area-width, 100vw);
    height: env(titlebar-area-height, 32px);
    -webkit-app-region: drag;
    z-index: 2147483647;
  }
  /* Without an Outlook top bar the page's scrollbar would run up under the
     window controls. Hide it; the page still scrolls with wheel, touchpad and
     keys. */
  html.${NO_TOP_BAR_CLASS} {
    scrollbar-width: none;
  }
`;

function isEnabled(config, platform = process.platform) {
  return config?.titleBar === "integrated" && platform !== "darwin";
}

function parseRgb(value) {
  const match = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?\s*\)/.exec(value || "");
  if (!match) return null;
  const alphaText = match[4] ?? "";
  let alpha = 1;
  if (alphaText) {
    alpha = alphaText.endsWith("%") ? Number.parseFloat(alphaText) / 100 : Number.parseFloat(alphaText);
  }
  return { r: Number(match[1]), g: Number(match[2]), b: Number(match[3]), a: alpha };
}

function toHex({ r, g, b }) {
  return `#${[r, g, b].map((n) => Math.min(255, n).toString(16).padStart(2, "0")).join("")}`;
}

// Used when the page has no readable background colour, such as the Microsoft
// sign-in pages, whose background is an image. Those pages follow the system
// colour scheme, so the controls do too.
function schemeFallback(matchMedia = globalThis.matchMedia) {
  try {
    return matchMedia?.(DARK_SCHEME).matches ? DARK_PAGE : LIGHT_PAGE;
  } catch {
    return LIGHT_PAGE;
  }
}

// Relative luminance (WCAG) to pick light or dark control glyphs.
function isDark({ r, g, b }) {
  const lin = (c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b) < 0.4;
}

// First opaque background walking up from the element.
function effectiveBackground(element) {
  for (let el = element; el; el = el.parentElement) {
    const rgb = parseRgb(getComputedStyle(el).backgroundColor);
    if (rgb && rgb.a > 0.5) return rgb;
  }
  return null;
}

// The top bar is the widest element, from a known top-bar control upwards,
// that is pinned to the top edge and no taller than a title bar (preferring
// one that paints a background, see paintsBackground). Wrappers
// between the control and the bar are often vertically centred, so they are
// skipped rather than ending the walk; it ends once an element is taller than
// a title bar.
function findTopBar() {
  let anchor = null;
  for (const selector of ANCHOR_SELECTORS) {
    anchor = document.querySelector(selector);
    if (anchor) break;
  }
  if (!anchor) return null;

  let bar = null;
  let paintedBar = null;
  for (let el = anchor; el && el !== document.body; el = el.parentElement) {
    const rect = el.getBoundingClientRect();
    if (rect.height > MAX_HEIGHT) break;
    if (rect.top <= 1 && rect.width >= window.innerWidth * 0.9 && rect.height >= MIN_HEIGHT) {
      bar = el;
      if (paintsBackground(el)) paintedBar = el;
    }
  }
  return paintedBar ?? bar;
}

// Outlook paints its top bar on an inner element inside a transparent,
// equally sized wrapper (#OwaTitleBar). Padding the wrapper to clear the
// window controls would leave that corner unpainted, so the page background
// showed through behind the controls as a lighter block. The painted element
// is the bar when there is one.
function paintsBackground(el) {
  const rgb = parseRgb(getComputedStyle(el).backgroundColor);
  return Boolean(rgb && rgb.a > 0.5);
}

// Outlook may support Window Controls Overlay itself (it ships as a PWA) and
// pad its bar's right edge by the controls' width. Padding again would leave a
// gap of twice that width, so the fallback padding is only used when it does
// not. The fallback padding sits on the bar itself, and that padding is ours,
// not Outlook's, so it does not count; otherwise the padding would be added
// and removed on every check (backported from teams-for-linux 4a1605a1).
function barReservesControls(bar, area) {
  if (bar.classList.contains(RESERVE_CLASS)) return false;
  const controlsWidth = window.innerWidth - (area.x + area.width);
  return Number.parseFloat(getComputedStyle(bar).paddingRight) >= controlsWidth - 1;
}

// Whether any of Outlook's own controls in the bar sit under the window controls.
function overlapsControls(bar, area) {
  const controlsLeft = area.x + area.width;
  for (const el of bar.querySelectorAll(INTERACTIVE_SELECTOR)) {
    const rect = el.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0 && rect.right > controlsLeft + 1 && rect.top < area.height) {
      return true;
    }
  }
  return false;
}

class IntegratedTitleBar {
  #ipcRenderer = null;
  #lastSent = "";
  #scheduled = false;
  #dragBar = null;
  #reservedBar = null;

  init(config, ipcRenderer) {
    if (!isEnabled(config) || !ipcRenderer) return;
    this.#ipcRenderer = ipcRenderer;

    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = STYLE;
    document.head.appendChild(style);

    const schedule = () => this.#schedule();
    new MutationObserver(schedule).observe(document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class", "data-theme"],
    });
    window.addEventListener("resize", schedule);
    globalThis.matchMedia?.(DARK_SCHEME)?.addEventListener?.("change", schedule);
    navigator.windowControlsOverlay?.addEventListener?.("geometrychange", schedule);
    this.#check();
  }

  #schedule() {
    if (this.#scheduled) return;
    this.#scheduled = true;
    setTimeout(() => {
      this.#scheduled = false;
      this.#check();
    }, CHECK_DELAY_MS);
  }

  #check() {
    try {
      const bar = findTopBar();
      this.#toggleDragStrip(!bar);
      this.#markDragBar(bar);
      if (bar) {
        this.#reserveSpace(bar);
      }
      this.#report(bar);
    } catch (err) {
      console.debug("[TITLE_BAR] Top bar check failed", { error: err.message });
    }
  }

  #toggleDragStrip(show) {
    document.documentElement.classList.toggle(NO_TOP_BAR_CLASS, show);
    const strip = document.getElementById(DRAG_STRIP_ID);
    if (show && !strip && document.body) {
      const div = document.createElement("div");
      div.id = DRAG_STRIP_ID;
      document.body.appendChild(div);
    } else if (!show && strip) {
      strip.remove();
    }
  }

  // Unlike Teams, Outlook does not mark its top bar as a drag region, so the
  // bar gets one here, with its interactive descendants carved out.
  #markDragBar(bar) {
    if (this.#dragBar === bar) return;
    this.#dragBar?.classList.remove(DRAG_CLASS);
    bar?.classList.add(DRAG_CLASS);
    this.#dragBar = bar;
  }

  #reserveSpace(bar) {
    const wco = navigator.windowControlsOverlay;
    const area = wco?.visible ? wco.getTitlebarAreaRect() : null;
    const outlookReserves = area !== null && barReservesControls(bar, area);
    if (this.#reservedBar && (this.#reservedBar !== bar || !area || outlookReserves)) {
      this.#reservedBar.classList.remove(RESERVE_CLASS);
      this.#reservedBar = null;
    }
    if (!area || outlookReserves || this.#reservedBar) return;
    if (overlapsControls(bar, area)) {
      bar.classList.add(RESERVE_CLASS);
      this.#reservedBar = bar;
    }
  }

  // Without a top bar (sign-in, loading) the controls keep the default height
  // and read the page background. The colour is sent so the main process can
  // derive a matching hover background. Outlook's top bar and the sign-in
  // pages paint images, which have no computed colour: then the system colour
  // scheme is only a placeholder and `sample` asks the main process to sample
  // the rendered page behind the controls instead.
  #report(bar) {
    const height = bar ? Math.round(bar.getBoundingClientRect().height) : DEFAULT_HEIGHT;
    if (height < MIN_HEIGHT || height > MAX_HEIGHT) return;
    const readable = effectiveBackground(bar ?? document.body);
    const bg = readable ?? schemeFallback();
    const update = {
      height,
      color: toHex(bg),
      symbolColor: isDark(bg) ? LIGHT_GLYPHS : DARK_GLYPHS,
      sample: !readable,
    };
    // A theme switch changes the image without changing anything above, so
    // while sampling, a page that keeps mutating is resampled at most once per
    // bucket instead of never.
    const bucket = readable ? 0 : Math.floor(Date.now() / RESAMPLE_INTERVAL_MS);
    const key = JSON.stringify({ ...update, bucket });
    if (key === this.#lastSent) return;
    this.#lastSent = key;
    this.#ipcRenderer.send(CHANNEL, update);
  }
}

module.exports = new IntegratedTitleBar();
module.exports._internals = { isEnabled, parseRgb, toHex, isDark, schemeFallback, barReservesControls, STYLE, ANCHOR_SELECTORS };
