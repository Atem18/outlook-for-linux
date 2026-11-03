"use strict";

// mailto: links received as launch arguments, turned into Outlook's compose
// deeplink. Pure, so it is unit tested without Electron. Modelled on
// teams-for-linux app/mainAppWindow/protocolLink.js.
//
// Never log the link or anything parsed from it: it carries addresses and
// message text. Log counts instead.

// Schemes the app asks the OS to route to it. Keep in step with
// `build.protocols.schemes` in package.json and MimeType in the Flatpak
// desktop entry.
const PROTOCOL_SCHEMES = ["mailto"];

const MAILTO_PREFIX = /^mailto:/i;

// Fields Outlook's /mail/deeplink/compose reads. Other RFC 6068 headers
// (in-reply-to, keywords, ...) have no equivalent there and are dropped.
const RECIPIENT_FIELDS = ["to", "cc", "bcc"];
const TEXT_FIELDS = ["subject", "body"];

function decode(text) {
  try {
    return decodeURIComponent(text);
  } catch {
    // A stray "%" is kept as typed rather than losing the whole link.
    return text;
  }
}

function splitAddresses(text) {
  return decode(text)
    .split(/[,;]/)
    .map((address) => address.trim())
    .filter(Boolean);
}

/**
 * Parses a mailto: URI (RFC 6068). Unlike a query string, "+" is a literal
 * plus (as in user+tag@example.com), never a space, so URLSearchParams would
 * get it wrong. Header names are case-insensitive; repeated recipient headers
 * add up, repeated text headers keep the first.
 * @param {string} uri
 * @returns {{to: string[], cc: string[], bcc: string[], subject: string, body: string}|null}
 *   null when the argument is not a mailto: link
 */
function parseMailto(uri) {
  if (typeof uri !== "string" || !MAILTO_PREFIX.test(uri)) return null;
  const rest = uri.replace(MAILTO_PREFIX, "");
  const queryStart = rest.indexOf("?");
  const path = queryStart === -1 ? rest : rest.slice(0, queryStart);
  const query = queryStart === -1 ? "" : rest.slice(queryStart + 1);

  const fields = { to: splitAddresses(path), cc: [], bcc: [], subject: "", body: "" };
  for (const pair of query.split("&")) {
    if (!pair) continue;
    const eq = pair.indexOf("=");
    const name = decode(eq === -1 ? pair : pair.slice(0, eq)).toLowerCase();
    const value = eq === -1 ? "" : pair.slice(eq + 1);
    if (RECIPIENT_FIELDS.includes(name)) {
      fields[name].push(...splitAddresses(value));
    } else if (TEXT_FIELDS.includes(name) && !fields[name]) {
      fields[name] = decode(value);
    }
  }
  return fields;
}

/**
 * The first mailto: link among command line arguments, or null.
 * @param {string[]} args
 * @returns {string|null}
 */
function findMailtoArg(args) {
  if (!Array.isArray(args)) return null;
  return args.find((arg) => typeof arg === "string" && MAILTO_PREFIX.test(arg)) ?? null;
}

/**
 * Outlook's compose deeplink for a parsed mailto: link, on the given mailbox
 * origin. The same path works for work (outlook.cloud.microsoft,
 * outlook.office.com) and personal (outlook.live.com) mailboxes.
 *
 * Built with encodeURIComponent, not URLSearchParams: the compose page does
 * not turn "+" back into a space, so URLSearchParams' "+" for spaces showed
 * up literally in the subject and body. Outlook currently ignores cc and bcc
 * here; they are passed anyway, and the counts are logged.
 * @param {ReturnType<typeof parseMailto>} fields
 * @param {string} origin - e.g. "https://outlook.live.com"
 * @returns {string}
 */
function composeUrl(fields, origin) {
  const url = new URL("/mail/deeplink/compose", origin);
  const params = [
    ...RECIPIENT_FIELDS.filter((name) => fields[name].length).map((name) => [name, fields[name].join(",")]),
    ...TEXT_FIELDS.filter((name) => fields[name]).map((name) => [name, fields[name]]),
  ];
  url.search = params.map(([name, value]) => `${name}=${encodeURIComponent(value)}`).join("&");
  return url.toString();
}

module.exports = { PROTOCOL_SCHEMES, parseMailto, findMailtoArg, composeUrl };
