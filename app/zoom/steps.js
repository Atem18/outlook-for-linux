"use strict";

// Zoom steps, as factors. Pure, so it is unit tested without Electron.
// Chrome's own presets, so Ctrl+Plus / Ctrl+Minus feel the same as in the
// browser.
const ZOOM_FACTORS = [0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

/**
 * The next zoom factor in a direction from the current one.
 * @param {number} current - current zoom factor
 * @param {1|-1|0} direction - in, out, or reset
 * @returns {number}
 */
function nextZoomFactor(current, direction) {
  if (direction === 0 || !Number.isFinite(current)) return 1;
  if (direction > 0) return ZOOM_FACTORS.find((f) => f > current + 0.001) ?? ZOOM_FACTORS.at(-1);
  return ZOOM_FACTORS.findLast((f) => f < current - 0.001) ?? ZOOM_FACTORS[0];
}

/**
 * A stored zoom factor, or 1 when it is missing or out of range.
 * @param {unknown} value
 * @returns {number}
 */
function validZoomFactor(value) {
  return typeof value === "number" && value >= ZOOM_FACTORS[0] && value <= ZOOM_FACTORS.at(-1) ? value : 1;
}

const ZOOM_IN_KEYS = new Set(["+", "="]);
const ZOOM_OUT_KEYS = new Set(["-", "_"]);

/**
 * The zoom a key press asks for, as Chrome reads it: Ctrl with + or = (the
 * same key on US layouts), - or _, 0, or the numpad's +, - and 0. `key` is
 * layout-aware; Ctrl+0 also matches by position because 0 needs Shift on
 * AZERTY. Alt and Meta are excluded so AltGr combinations stay with the page.
 * @param {{type?: string, key?: string, code?: string, control?: boolean, alt?: boolean, meta?: boolean}} input -
 *   Electron's before-input-event input
 * @returns {1|-1|0|null} in, out, reset, or not a zoom key
 */
function zoomDirectionForInput(input) {
  if (input?.type !== "keyDown" || !input.control || input.alt || input.meta) return null;
  if (ZOOM_IN_KEYS.has(input.key) || input.code === "NumpadAdd") return 1;
  if (ZOOM_OUT_KEYS.has(input.key) || input.code === "NumpadSubtract") return -1;
  if (input.key === "0" || input.code === "Digit0" || input.code === "Numpad0") return 0;
  return null;
}

module.exports = { ZOOM_FACTORS, nextZoomFactor, validZoomFactor, zoomDirectionForInput };
