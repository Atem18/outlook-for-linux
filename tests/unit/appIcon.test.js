const { test } = require("node:test");
const assert = require("node:assert/strict");
const { customIconPath, windowIconResize } = require("../../app/appIcon/iconPath");

test("the appIcon option names the custom icon; blank or non-string means the bundled one", () => {
  assert.equal(customIconPath(" /a/option.png "), "/a/option.png");
  assert.equal(customIconPath(""), null);
  assert.equal(customIconPath("   "), null);
  assert.equal(customIconPath(undefined), null);
  assert.equal(customIconPath(42), null);
});

test("window icons above 128px shrink on the longer side; smaller ones stay", () => {
  assert.deepEqual(windowIconResize({ width: 512, height: 512 }), { width: 128 });
  assert.deepEqual(windowIconResize({ width: 200, height: 400 }), { height: 128 });
  assert.equal(windowIconResize({ width: 128, height: 64 }), null);
});
