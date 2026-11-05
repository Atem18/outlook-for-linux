"use strict";

// Pure helpers for the unread badge, shared by the tray (main process) and
// unit tests.

const APP_NAME = "Outlook for Linux";

// A rendered 64x64 PNG badge is a few KB; anything far larger is not one.
const MAX_BADGE_DATA_URL_LENGTH = 200_000;
const PNG_DATA_URL_PREFIX = "data:image/png;base64,";

function trayTooltip(count) {
  return count > 0 ? `${APP_NAME} (${count})` : APP_NAME;
}

function isBadgeDataUrl(value) {
  return (
    typeof value === "string" &&
    value.startsWith(PNG_DATA_URL_PREFIX) &&
    value.length <= MAX_BADGE_DATA_URL_LENGTH
  );
}

/**
 * Validate an unread update from the renderer.
 * @param {unknown} payload
 * @returns {{count: number, badge: string|null} | null}
 */
function sanitizeUnreadUpdate(payload) {
  if (!payload || typeof payload !== "object") return null;
  const { count, badge } = payload;
  if (!Number.isInteger(count) || count < 0 || count > 1_000_000) return null;
  return { count, badge: isBadgeDataUrl(badge) ? badge : null };
}

/**
 * Draws the badge over the icon: "source over" on premultiplied-alpha pixels,
 * which is what nativeImage.toBitmap() returns. Channel order does not matter
 * as long as both buffers come from toBitmap() on the same machine.
 * @param {Buffer} base - icon pixels, 4 bytes each
 * @param {Buffer} badge - badge pixels, same size
 * @returns {Buffer|null} the composited pixels, or null on a size mismatch
 */
function compositeBadge(base, badge) {
  if (!base || !badge || base.length !== badge.length || base.length % 4 !== 0) return null;
  const out = Buffer.alloc(base.length);
  for (let i = 0; i < base.length; i += 4) {
    const inverseAlpha = 255 - badge[i + 3];
    for (let c = 0; c < 4; c++) {
      out[i + c] = badge[i + c] + Math.round((base[i + c] * inverseAlpha) / 255);
    }
  }
  return out;
}

/**
 * The sum of the accounts' unread counts, ignoring anything that is not a
 * non-negative integer.
 * @param {Iterable<unknown>} counts
 * @returns {number}
 */
function unreadTotal(counts) {
  let total = 0;
  for (const count of counts) {
    if (Number.isInteger(count) && count > 0) total += count;
  }
  return total;
}

module.exports = { trayTooltip, isBadgeDataUrl, sanitizeUnreadUpdate, compositeBadge, unreadTotal };
