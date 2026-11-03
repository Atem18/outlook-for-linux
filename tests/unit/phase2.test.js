const { test } = require("node:test");
const assert = require("node:assert/strict");
const { nextZoomFactor, validZoomFactor, ZOOM_FACTORS } = require("../../app/zoom/steps");

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

test("a stored zoom outside the presets' range falls back to 100%", () => {
  assert.equal(validZoomFactor(1.25), 1.25);
  assert.equal(validZoomFactor(undefined), 1);
  assert.equal(validZoomFactor("2"), 1);
  assert.equal(validZoomFactor(10), 1);
});
