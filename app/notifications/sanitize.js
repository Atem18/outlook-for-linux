"use strict";

// Validation for notification requests coming from the page. Outlook's
// notifications carry the sender and subject, so lengths are capped and
// nothing here is ever logged.

const MAX_TITLE = 256;
const MAX_BODY = 1024;
const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;

function clip(value, max) {
  return typeof value === "string" ? value.slice(0, max) : "";
}

/**
 * @param {unknown} payload
 * @returns {{id: string, title: string, body: string, silent: boolean, requireInteraction: boolean} | null}
 */
function sanitizeNotification(payload) {
  if (!payload || typeof payload !== "object") return null;
  const { id, title, body, silent, requireInteraction } = payload;
  if (typeof id !== "string" || !ID_RE.test(id)) return null;
  const cleanTitle = clip(title, MAX_TITLE);
  if (!cleanTitle) return null;
  return {
    id,
    title: cleanTitle,
    body: clip(body, MAX_BODY),
    silent: silent === true,
    requireInteraction: requireInteraction === true,
  };
}

function isNotificationId(value) {
  return typeof value === "string" && ID_RE.test(value);
}

module.exports = { sanitizeNotification, isNotificationId };
