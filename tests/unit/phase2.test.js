const { test } = require("node:test");
const assert = require("node:assert/strict");
const { nextZoomFactor, validZoomFactor, zoomDirectionForInput, ZOOM_FACTORS } = require("../../app/zoom/steps");

test("zoom steps through Chrome's presets and stops at the ends", () => {
  assert.equal(nextZoomFactor(1, 1), 1.1);
  assert.equal(nextZoomFactor(1, -1), 0.9);
  assert.equal(nextZoomFactor(1.17, 1), 1.25);
  assert.equal(nextZoomFactor(1.17, -1), 1.1);
  assert.equal(nextZoomFactor(ZOOM_FACTORS.at(-1), 1), ZOOM_FACTORS.at(-1));
  assert.equal(nextZoomFactor(ZOOM_FACTORS[0], -1), ZOOM_FACTORS[0]);
  assert.equal(nextZoomFactor(2.5, 0), 1);
  assert.equal(nextZoomFactor(Number.NaN, 1), 1);
});

test("Ctrl with +, =, -, _, 0 or the numpad keys zooms; other keys do not", () => {
  const key = (props) => zoomDirectionForInput({ type: "keyDown", control: true, ...props });
  assert.equal(key({ key: "=", code: "Equal" }), 1);
  assert.equal(key({ key: "+", code: "Equal", shift: true }), 1);
  assert.equal(key({ key: "+", code: "NumpadAdd" }), 1);
  assert.equal(key({ key: "-", code: "Minus" }), -1);
  assert.equal(key({ key: "_", code: "Minus", shift: true }), -1);
  assert.equal(key({ key: "-", code: "NumpadSubtract" }), -1);
  assert.equal(key({ key: "0", code: "Digit0" }), 0);
  assert.equal(key({ key: "Insert", code: "Numpad0" }), 0);
  // AZERTY: the 0 key types "à" without Shift.
  assert.equal(key({ key: "à", code: "Digit0" }), 0);
  // AZERTY: "-" sits on the 6 key.
  assert.equal(key({ key: "-", code: "Digit6" }), -1);
  assert.equal(key({ key: "r", code: "KeyR" }), null);
  assert.equal(key({ key: "=", code: "Equal", alt: true }), null);
  assert.equal(key({ key: "=", code: "Equal", meta: true }), null);
  assert.equal(zoomDirectionForInput({ type: "keyDown", key: "=", code: "Equal", control: false }), null);
  assert.equal(zoomDirectionForInput({ type: "keyUp", key: "=", code: "Equal", control: true }), null);
  assert.equal(zoomDirectionForInput(undefined), null);
});

test("a stored zoom outside the presets' range falls back to 100%", () => {
  assert.equal(validZoomFactor(1.25), 1.25);
  assert.equal(validZoomFactor(undefined), 1);
  assert.equal(validZoomFactor("2"), 1);
  assert.equal(validZoomFactor(10), 1);
});
