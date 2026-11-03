"use strict";

// Which icon applies, and at what size for a window. Pure, so it is unit
// tested without Electron.

/**
 * @param {unknown} configured - the appIcon option
 * @returns {string|null} the custom icon's path, or null for the bundled icon
 */
function customIconPath(configured) {
  return typeof configured === "string" && configured.trim() ? configured.trim() : null;
}

/**
 * Size for a window icon: the desktop never shows one above 128px, and a big
 * custom PNG would otherwise be kept and sent at full size. Keeps the aspect
 * ratio on the longer side.
 * @param {{width: number, height: number}} size
 * @returns {{width: number}|{height: number}|null} null to keep the image as is
 */
function windowIconResize({ width, height }) {
  if (Math.max(width, height) <= 128) return null;
  return width >= height ? { width: 128 } : { height: 128 };
}

module.exports = { customIconPath, windowIconResize };
