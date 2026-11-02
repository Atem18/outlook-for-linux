"use strict";

// Pure helpers for reading the Inbox unread count from Outlook's own API
// (/owa/service.svc, the JSON flavour of EWS that the Outlook web app calls).

// Origins whose service.svc calls may lend their token, and the only origins
// it is ever sent back to. Same list as the remembered mailbox (startUrl.js).
const API_ORIGINS = new Set([
  "https://outlook.live.com",
  "https://outlook.cloud.microsoft",
  "https://outlook.office.com",
  "https://outlook.office365.com",
]);

/**
 * The mailbox origin of an Outlook API request whose Authorization header can
 * be reused, or null.
 * @param {string} url
 * @returns {string|null}
 */
function apiOriginOf(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return null;
  }
  if (!API_ORIGINS.has(parsed.origin)) return null;
  return /^\/owa\/service\.svc$/i.test(parsed.pathname) ? parsed.origin : null;
}

// Header names are case-insensitive; Electron passes them as sent.
function bearerFrom(headers) {
  for (const [name, value] of Object.entries(headers ?? {})) {
    if (name.toLowerCase() === "authorization" && typeof value === "string" && /^Bearer \S+$/.test(value)) {
      return value;
    }
  }
  return null;
}

function getFolderUrl(origin) {
  return `${origin}/owa/service.svc?action=GetFolder&app=Mail`;
}

// GetFolder on the well-known Inbox, asking only for its unread count.
function getFolderBody() {
  return JSON.stringify({
    __type: "GetFolderJsonRequest:#Exchange",
    Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "V2018_01_08" },
    Body: {
      __type: "GetFolderRequest:#Exchange",
      FolderShape: {
        __type: "FolderResponseShape:#Exchange",
        BaseShape: "IdOnly",
        AdditionalProperties: [{ __type: "PropertyUri:#Exchange", FieldURI: "folder:UnreadCount" }],
      },
      FolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "inbox" }],
    },
  });
}

/**
 * @param {unknown} json - a GetFolder response
 * @returns {number|null} the Inbox unread count, or null if absent
 */
function parseUnreadCount(json) {
  const count = json?.Body?.ResponseMessages?.Items?.[0]?.Folders?.[0]?.UnreadCount;
  return Number.isInteger(count) && count >= 0 ? count : null;
}

module.exports = { apiOriginOf, bearerFrom, getFolderUrl, getFolderBody, parseUnreadCount };
