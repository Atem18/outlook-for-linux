/**
 * Integrated title bar for pop-out windows (the mailto: compose window).
 *
 * The main window's controls sit in Outlook's top bar (see
 * integratedTitleBar.js), but a pop-out has none: its ribbon tabs start at
 * the top edge. Outlook for Windows gives its pop-outs a title bar of their
 * own above the ribbon, so this does the same: a strip showing the window
 * title, in the main window's header colour, that the window is dragged by
 * and that the window controls sit in. The page moves down to make room.
 */

const { toHex, isDark, schemeFallback } = require("./integratedTitleBar")._internals;

const CHANNEL = "title-bar-overlay-update";
const STRIP_ID = "ofl-pop-out-title-bar";
const ACTIVE_CLASS = "ofl-pop-out";
const HEIGHT = 32;
const LIGHT_GLYPHS = "#ffffff";
const DARK_GLYPHS = "#242424";
const HEX_COLOR_RE = /^#[0-9a-f]{6}$/i;

const STYLE = `
  html.${ACTIVE_CLASS} {
    box-sizing: border-box !important;
    height: 100vh !important;
    padding-top: ${HEIGHT}px !important;
  }
  html.${ACTIVE_CLASS} > body {
    height: 100% !important;
  }
  #${STRIP_ID} {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    height: ${HEIGHT}px;
    box-sizing: border-box;
    padding: 0 calc(100vw - env(titlebar-area-x, 0px) - env(titlebar-area-width, 100vw) + 12px) 0 12px;
    display: flex;
    align-items: center;
    font: 12px "Segoe UI", system-ui, sans-serif;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    -webkit-app-region: drag;
    user-select: none;
    z-index: 2147483647;
  }
`;

// The strip's colour: the main window's header colour, which the main
// process samples and passes in (and sends again when it changes). Before that (a
// mailto: link that started the app), the system colour scheme.
function stripColors(headerColor) {
  const rgb = HEX_COLOR_RE.test(headerColor ?? "")
    ? { r: parseInt(headerColor.slice(1, 3), 16), g: parseInt(headerColor.slice(3, 5), 16), b: parseInt(headerColor.slice(5, 7), 16) }
    : schemeFallback();
  return { color: toHex(rgb), symbolColor: isDark(rgb) ? LIGHT_GLYPHS : DARK_GLYPHS };
}

class PopOutTitleBar {
  #ipcRenderer = null;
  #headerColor = null;
  #lastSent = "";

  /**
   * @param {{titleBar?: string, popOutTitleBar?: {color?: string}}} config
   * @param {Electron.IpcRenderer} ipcRenderer
   */
  init(config, ipcRenderer) {
    if (!config?.popOutTitleBar || !ipcRenderer) return;
    this.#ipcRenderer = ipcRenderer;
    this.#headerColor = config.popOutTitleBar.color ?? null;
    ipcRenderer.on("pop-out-header-color", (_event, color) => {
      if (!HEX_COLOR_RE.test(color ?? "")) return;
      this.#headerColor = color;
      this.#update();
    });

    const style = document.createElement("style");
    style.textContent = STYLE;
    document.head.appendChild(style);
    document.documentElement.classList.add(ACTIVE_CLASS);

    // The title changes as Outlook loads and as the subject is typed.
    new MutationObserver(() => this.#update()).observe(document.head, {
      childList: true,
      subtree: true,
      characterData: true,
    });
    this.#update();
  }

  #update() {
    try {
      let strip = document.getElementById(STRIP_ID);
      if (!strip && document.body) {
        strip = document.createElement("div");
        strip.id = STRIP_ID;
        document.body.appendChild(strip);
      }
      if (!strip) return;
      const { color, symbolColor } = stripColors(this.#headerColor);
      strip.textContent = document.title;
      strip.style.background = color;
      strip.style.color = symbolColor;
      const update = { height: HEIGHT, color, symbolColor, sample: false };
      const key = JSON.stringify(update);
      if (key === this.#lastSent) return;
      this.#lastSent = key;
      this.#ipcRenderer.send(CHANNEL, update);
    } catch (err) {
      console.debug("[TITLE_BAR] Pop-out title bar update failed", { error: err.message });
    }
  }
}

module.exports = new PopOutTitleBar();
module.exports._internals = { HEIGHT, STYLE, stripColors };
