"use strict";

// .eml files (RFC 5322 messages) opened from the file manager, the command
// line or File → Open Message…, turned into what the viewer window shows.
// Electron-free, so it is unit tested with node:test. Parsing is postal-mime's.
//
// Never log anything read from a message (addresses, subject, body, file
// names): log counts and sizes instead.

const path = require("node:path");
const { fileURLToPath } = require("node:url");
const PostalMime = require("postal-mime").default;

const EML_EXTENSION = /\.eml$/i;

// Larger files are refused before they are read: a message this size is a
// mistake or an attack, and the whole file is held in memory.
const MAX_EML_BYTES = 100 * 1024 * 1024;

// The quoted original in a reply travels in the compose deeplink's URL.
const MAX_REPLY_QUOTE_CHARS = 10_000;

/**
 * The .eml files among command line arguments, as absolute paths. A second
 * launch's relative paths are relative to its own working directory, not
 * this process's. Desktop entries with %U pass file:// URIs.
 * @param {string[]} args
 * @param {string} workingDirectory
 * @returns {string[]}
 */
function findEmlArgs(args, workingDirectory) {
  if (!Array.isArray(args)) return [];
  const files = [];
  for (const arg of args) {
    if (typeof arg !== "string" || arg.startsWith("-") || !EML_EXTENSION.test(arg)) continue;
    let file = arg;
    if (/^file:/i.test(arg)) {
      try {
        file = fileURLToPath(arg);
      } catch {
        continue;
      }
    } else if (/^[a-z][a-z0-9+.-]*:/i.test(arg)) {
      // Another scheme (https:, mailto:, ...) is not a local file.
      continue;
    }
    const absolute = path.resolve(workingDirectory, file);
    if (!files.includes(absolute)) files.push(absolute);
  }
  return files;
}

function formatMailbox({ name, address }) {
  if (!address) return name || "";
  return name && name !== address ? `${name} <${address}>` : address;
}

// Groups ("Team: a@x, b@y;") are flattened: the viewer lists mailboxes.
function mailboxes(addresses) {
  if (!addresses) return [];
  const list = Array.isArray(addresses) ? addresses : [addresses];
  return list.flatMap((entry) => (entry.group ? entry.group : [entry])).filter((entry) => entry.address || entry.name);
}

function contentIdOf(attachment) {
  return attachment.contentId ? attachment.contentId.replace(/^<|>$/g, "").trim() : "";
}

function toBuffer(content) {
  if (typeof content === "string") return Buffer.from(content, "utf8");
  if (content instanceof ArrayBuffer) return Buffer.from(content);
  return Buffer.from(content.buffer, content.byteOffset, content.byteLength);
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Replaces cid: references to the message's own images with data: URIs, so
 * the viewer needs no network and no file access for them.
 * @returns {{html: string, inlined: Set<number>}} indexes of the attachments used
 */
function inlineCidImages(html, attachments) {
  const inlined = new Set();
  let result = html;
  attachments.forEach((attachment, index) => {
    const id = contentIdOf(attachment);
    if (!id || !attachment.mimeType.startsWith("image/")) return;
    const reference = new RegExp(`cid:${escapeRegExp(id)}(?=["'\\s)>]|$)`, "gi");
    if (!reference.test(result)) return;
    const dataUri = `data:${attachment.mimeType};base64,${toBuffer(attachment.content).toString("base64")}`;
    result = result.replace(reference, () => dataUri);
    inlined.add(index);
  });
  return { html: result, inlined };
}

function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
}

// A plain-text body as HTML: escaped, line breaks kept, web links clickable.
function textToHtml(text) {
  const linked = escapeHtml(text).replace(/\bhttps?:\/\/[^\s<>"']+[^\s<>"'.,;:!?)\]]/g, (url) => `<a href="${url}">${url}</a>`);
  return `<div style="white-space: pre-wrap; font-family: monospace; font-size: 13px">${linked}</div>`;
}

// Anything that would load from the network when the HTML renders.
const REMOTE_REFERENCE = /(?:\b(?:src|srcset|background|poster)\s*=\s*["']?\s*(?:https?:)?\/\/)|(?:url\(\s*["']?\s*(?:https?:)?\/\/)|(?:@import\s+["']?(?:https?:)?\/\/)/i;

/**
 * The parts of a parsed message the viewer shows.
 * @param {import("postal-mime").Email} email
 */
function toViewModel(email) {
  const attachments = email.attachments ?? [];
  let html;
  let inlined = new Set();
  if (email.html) {
    ({ html, inlined } = inlineCidImages(email.html, attachments));
  } else {
    html = textToHtml(email.text ?? "");
  }
  const date = email.date ? new Date(email.date) : null;
  return {
    subject: email.subject ?? "",
    from: mailboxes(email.from).map(formatMailbox),
    to: mailboxes(email.to).map(formatMailbox),
    cc: mailboxes(email.cc).map(formatMailbox),
    // Kept as an ISO string; the viewer formats it in the user's locale.
    date: date && !Number.isNaN(date.getTime()) ? date.toISOString() : null,
    html,
    hasRemoteContent: REMOTE_REFERENCE.test(html),
    // Indexes point into the parsed message's attachments, for saving.
    attachments: attachments
      .map((attachment, index) => ({ attachment, index }))
      .filter(({ index }) => !inlined.has(index))
      .map(({ attachment, index }) => ({
        index,
        filename: attachmentFileName(attachment, index),
        mimeType: attachment.mimeType,
        size: toBuffer(attachment.content).byteLength,
      })),
  };
}

/**
 * A file name to save an attachment under: the message's own, without any
 * directory part, or a generic one.
 */
function attachmentFileName(attachment, index) {
  const name = path.basename((attachment.filename ?? "").replaceAll("\\", "/")).replace(/\p{Cc}/gu, "").trim();
  if (name && name !== "." && name !== "..") return name;
  if (attachment.mimeType === "message/rfc822") return `message-${index + 1}.eml`;
  return `attachment-${index + 1}${extensionFor(attachment.mimeType)}`;
}

// Extensions for unnamed parts, so a saved file opens in the right app.
// Bounce and abuse reports (delivery-status, feedback-report) are text.
const EXTENSIONS = {
  "text/html": ".html",
  "text/calendar": ".ics",
  "text/csv": ".csv",
  "text/rfc822-headers": ".txt",
  "message/delivery-status": ".txt",
  "message/disposition-notification": ".txt",
  "message/feedback-report": ".txt",
  "application/pdf": ".pdf",
  "application/zip": ".zip",
  "application/gzip": ".gz",
  "application/x-gzip": ".gz",
  "application/json": ".json",
  "application/pgp-signature": ".asc",
  "application/pkcs7-signature": ".p7s",
  "application/x-pkcs7-signature": ".p7s",
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/gif": ".gif",
  "image/webp": ".webp",
  "image/svg+xml": ".svg",
};

function extensionFor(mimeType) {
  if (EXTENSIONS[mimeType]) return EXTENSIONS[mimeType];
  return mimeType?.startsWith("text/") ? ".txt" : "";
}

/**
 * The HTML for the viewer's sandboxed iframe (srcdoc). The iframe runs no
 * scripts and submits no forms; this CSP also keeps the message from
 * loading anything from the network (tracking pixels) unless the user asks
 * for remote content. Links open in a new window, which the main process
 * hands to the browser.
 * @param {string} html - the message body as HTML
 * @param {{remoteContent?: boolean}} options
 */
function messageDocument(html, { remoteContent = false } = {}) {
  const remote = remoteContent ? " https: http:" : "";
  const csp = [
    "default-src 'none'",
    `img-src data:${remote}`,
    `style-src 'unsafe-inline'${remote}`,
    `font-src data:${remote}`,
    `media-src data:${remote}`,
    "form-action 'none'",
  ].join("; ");
  return (
    "<!doctype html><html><head>" +
    '<meta charset="utf-8">' +
    `<meta http-equiv="Content-Security-Policy" content="${csp}">` +
    '<base target="_blank">' +
    "<style>body { margin: 16px; background: #fff; color: #000; font-family: sans-serif; overflow-wrap: anywhere; } img { max-width: 100%; height: auto; }</style>" +
    "</head><body>" +
    html +
    "</body></html>"
  );
}

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };

// A rough plain-text version of an HTML body, for quoting in a reply when
// the message has no text part. It only feeds the compose deeplink's text.
function htmlToText(html) {
  return html
    .replace(/<(head|style|script|title)\b[\s\S]*?<\/\1\s*>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|tr|h[1-6]|blockquote|table)\s*>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (entity, code) => {
      if (code[0] === "#") {
        const point = code[1].toLowerCase() === "x" ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
        return point > 0 && point <= 0x10ffff ? String.fromCodePoint(point) : entity;
      }
      return ENTITIES[code.toLowerCase()] ?? entity;
    })
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * The compose fields for a reply: to the Reply-To (else the sender), "Re:"
 * once, and the original quoted below. Same shape as parseMailto's result.
 * @param {import("postal-mime").Email} email
 * @param {string} [dateText] - the original's date as the user reads it
 */
function replyFields(email, dateText) {
  const replyTo = mailboxes(email.replyTo);
  const recipients = (replyTo.length ? replyTo : mailboxes(email.from)).map((entry) => entry.address).filter(Boolean);
  const subject = email.subject ?? "";
  const sender = mailboxes(email.from).map(formatMailbox).join(", ");
  let original = (email.text ?? (email.html ? htmlToText(email.html) : "")).trim();
  if (original.length > MAX_REPLY_QUOTE_CHARS) original = `${original.slice(0, MAX_REPLY_QUOTE_CHARS)}\n[…]`;
  const intro = [dateText ? `On ${dateText}` : "", sender ? `${sender} wrote:` : "wrote:"].filter(Boolean).join(", ");
  const quoted = original ? original.split(/\r?\n/).map((line) => `> ${line}`).join("\n") : "";
  return {
    to: recipients,
    cc: [],
    bcc: [],
    subject: /^re:/i.test(subject.trim()) ? subject.trim() : `Re: ${subject.trim()}`.trim(),
    body: quoted ? `\n\n${intro}\n${quoted}` : "",
  };
}

/**
 * Parses a raw message.
 * @param {Buffer|Uint8Array|string} raw
 * @returns {Promise<import("postal-mime").Email>}
 */
function parseEml(raw) {
  return PostalMime.parse(raw, { attachmentEncoding: "arraybuffer" });
}

module.exports = {
  MAX_EML_BYTES,
  findEmlArgs,
  parseEml,
  toViewModel,
  inlineCidImages,
  messageDocument,
  replyFields,
  htmlToText,
  attachmentFileName,
  toBuffer,
};
