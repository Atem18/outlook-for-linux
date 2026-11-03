const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

const {
  isIntegratedTitleBar,
  titleBarWindowOptions,
  sanitizeOverlayUpdate,
  overlayColorsFromPixel,
  sampleRect,
  DEFAULT_HEIGHT,
} = require("../../app/mainAppWindow/titleBar");
const renderer = require("../../app/browser/tools/integratedTitleBar");
const options = require("../../app/config/options");

const integrated = { titleBar: "integrated" };

describe("isIntegratedTitleBar", () => {
  it("is on for \"integrated\" on Linux and Windows", () => {
    assert.equal(isIntegratedTitleBar(integrated, "linux"), true);
    assert.equal(isIntegratedTitleBar(integrated, "win32"), true);
  });

  it("is on by default, off when set to system and on macOS", () => {
    assert.equal(options.titleBar.default, "integrated");
    assert.equal(isIntegratedTitleBar({ titleBar: "system" }, "linux"), false);
    assert.equal(isIntegratedTitleBar({}, "linux"), false);
    assert.equal(isIntegratedTitleBar(integrated, "darwin"), false);
  });
});

describe("titleBarWindowOptions", () => {
  it("adds nothing when the integrated title bar is off", () => {
    assert.deepEqual(titleBarWindowOptions({ titleBar: "system" }, false, "linux"), {});
  });

  it("hides the title bar and draws the controls on a transparent overlay", () => {
    const dark = titleBarWindowOptions(integrated, true, "linux");
    assert.equal(dark.titleBarStyle, "hidden");
    assert.equal(dark.titleBarOverlay.height, DEFAULT_HEIGHT);
    assert.equal(dark.titleBarOverlay.color, "#00000000");
    assert.equal(dark.titleBarOverlay.symbolColor, "#ffffff");
    // A light page needs a light (if transparent) base colour, or the hover
    // background derived from it turns dark under dark glyphs.
    const light = titleBarWindowOptions(integrated, false, "linux");
    assert.equal(light.titleBarOverlay.color, "#ffffff00");
    assert.notEqual(light.titleBarOverlay.symbolColor, "#ffffff");
  });
});

describe("sanitizeOverlayUpdate", () => {
  it("accepts a well-formed update and makes the page colour transparent", () => {
    assert.deepEqual(
      sanitizeOverlayUpdate({ height: 48, color: "#ebebeb", symbolColor: "#242424", extra: 1 }),
      { height: 48, color: "#ebebeb00", symbolColor: "#242424", sample: false }
    );
  });

  it("passes the sample request through only when it is exactly true", () => {
    const base = { height: 48, color: "#ffffff", symbolColor: "#242424" };
    assert.equal(sanitizeOverlayUpdate({ ...base, sample: true }).sample, true);
    assert.equal(sanitizeOverlayUpdate({ ...base, sample: "yes" }).sample, false);
  });

  for (const [label, payload] of [
    ["null", null],
    ["a string", "x"],
    ["a fractional height", { height: 48.5, color: "#000000", symbolColor: "#ffffff" }],
    ["a too-small height", { height: 4, color: "#000000", symbolColor: "#ffffff" }],
    ["a too-large height", { height: 500, color: "#000000", symbolColor: "#ffffff" }],
    ["a missing page colour", { height: 48, symbolColor: "#ffffff" }],
    ["a page colour with its own alpha", { height: 48, color: "#ffffffff", symbolColor: "#242424" }],
    ["a CSS expression colour", { height: 48, color: "#000000", symbolColor: "url(x)" }],
    ["a short hex symbol colour", { height: 48, color: "#000000", symbolColor: "#fff" }],
  ]) {
    it(`rejects ${label}`, () => {
      assert.equal(sanitizeOverlayUpdate(payload), null);
    });
  }
});

describe("page sampling", () => {
  it("reads a BGRA pixel and picks contrasting glyphs", () => {
    // The light blue-green Outlook theme: dark glyphs, despite a dark system scheme.
    assert.deepEqual(overlayColorsFromPixel(Buffer.from([0xe0, 0xd8, 0xc8, 0xff])), {
      color: "#c8d8e000",
      symbolColor: "#242424",
    });
    assert.equal(overlayColorsFromPixel(Buffer.from([0x20, 0x20, 0x20, 0xff])).symbolColor, "#ffffff");
    assert.equal(overlayColorsFromPixel(Buffer.alloc(0)), null);
  });

  it("samples the right end of the top bar, clamped to the window", () => {
    assert.deepEqual(sampleRect(1000, 48), { x: 840, y: 0, width: 160, height: 48 });
    assert.deepEqual(sampleRect(100, 48), { x: 0, y: 0, width: 100, height: 48 });
  });
});

describe("integratedTitleBar renderer helpers", () => {
  const { isEnabled, parseRgb, toHex, isDark, schemeFallback, STYLE } = renderer._internals;

  it("matches the main-process enable rule", () => {
    assert.equal(isEnabled(integrated, "linux"), true);
    assert.equal(isEnabled(integrated, "darwin"), false);
    assert.equal(isEnabled({ titleBar: "system" }, "linux"), false);
  });

  it("parses computed colours, including alpha", () => {
    assert.deepEqual(parseRgb("rgb(31, 31, 31)"), { r: 31, g: 31, b: 31, a: 1 });
    assert.deepEqual(parseRgb("rgba(0, 0, 0, 0)"), { r: 0, g: 0, b: 0, a: 0 });
    assert.equal(parseRgb("transparent"), null);
  });

  it("converts to hex and picks glyph contrast from the background", () => {
    assert.equal(toHex({ r: 31, g: 0, b: 255 }), "#1f00ff");
    assert.equal(isDark({ r: 31, g: 31, b: 31 }), true);
    assert.equal(isDark({ r: 235, g: 235, b: 235 }), false);
  });

  it("follows the system colour scheme when the page colour is unreadable", () => {
    const media = (dark) => (query) => ({ matches: dark && query === "(prefers-color-scheme: dark)" });
    assert.equal(isDark(schemeFallback(media(true))), true);
    assert.equal(isDark(schemeFallback(media(false))), false);
    assert.equal(isDark(schemeFallback(() => { throw new Error("no media"); })), false);
  });

  it("makes the top bar draggable but keeps its controls clickable", () => {
    assert.match(STYLE, /\.ofl-title-bar-drag \{\s*-webkit-app-region: drag;/);
    assert.match(STYLE, /\.ofl-title-bar-drag :is\(button, a, input[^)]*\) \{\s*-webkit-app-region: no-drag;/);
  });

  // Backported from teams-for-linux 4a1605a1.
  it("does not take its own fallback padding for Outlook reserving the controls", () => {
    const { barReservesControls } = renderer._internals;
    const saved = { window: globalThis.window, getComputedStyle: globalThis.getComputedStyle };
    let ours = false;
    const bar = { classList: { contains: (name) => ours && name === "ofl-wco-reserve" } };
    globalThis.window = { innerWidth: 800 };
    globalThis.getComputedStyle = () => ({ paddingRight: "72px" });
    try {
      const area = { x: 0, width: 728 };
      assert.equal(barReservesControls(bar, area), true, "Outlook padding counts");
      ours = true;
      assert.equal(barReservesControls(bar, area), false, "our own padding does not");
    } finally {
      Object.assign(globalThis, saved);
    }
  });

  it("does nothing when disabled", () => {
    assert.doesNotThrow(() => renderer.init({ titleBar: "system" }, {}));
  });
});

it("pop-out title bar takes the main window's header colour", () => {
  const { stripColors, HEIGHT } = require("../../app/browser/tools/popOutTitleBar")._internals;
  assert.deepEqual(stripColors("#0e0e0e"), { color: "#0e0e0e", symbolColor: "#ffffff" });
  assert.deepEqual(stripColors("#f5f5f5"), { color: "#f5f5f5", symbolColor: "#242424" });
  // No header colour yet: the system scheme decides.
  assert.match(stripColors(null).color, /^#(000000|ffffff)$/);
  assert.equal(stripColors("not a colour").color.length, 7);
  assert.ok(HEIGHT >= 24 && HEIGHT <= 80);
});
