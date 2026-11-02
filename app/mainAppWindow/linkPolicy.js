"use strict";

// Decides where a URL belongs: inside the app (Outlook itself and the
// Microsoft sign-in pages it relies on) or in the user's browser (everything
// else, including links inside emails). Pure, so it is unit tested without
// Electron.

// Hosts that render Outlook or complete its sign-in. A suffix match, so
// "office.com" also covers "outlook.office.com" and "www.office.com".
const IN_APP_HOST_SUFFIXES = [
  "outlook.cloud.microsoft",
  "outlook.office.com",
  "outlook.office365.com",
  "outlook.live.com",
  "login.microsoftonline.com",
  "login.microsoft.com",
  "login.live.com",
  "account.live.com",
  "account.microsoft.com",
  "msauth.net",
  "msftauth.net",
];

// Subdomains of the hosts above that must still open externally. Defender
// Safe Links wraps every link in an email as
// https://<region>.safelinks.protection.outlook.com/?url=..., which would
// otherwise match outlook.com-style rules and load the target inside the app.
const EXTERNAL_HOST_SUFFIXES = [
  "safelinks.protection.outlook.com",
  "statics.teams.cdn.office.net",
];

function hostMatches(host, suffix) {
  return host === suffix || host.endsWith(`.${suffix}`);
}

function parseUrl(url) {
  try {
    return new URL(url);
  } catch {
    return null;
  }
}

/**
 * @param {string} url
 * @returns {boolean} true when the URL should load inside the app
 */
function isInAppUrl(url) {
  const parsed = parseUrl(url);
  if (parsed?.protocol !== "https:") return false;
  const host = parsed.hostname.toLowerCase();
  if (EXTERNAL_HOST_SUFFIXES.some((s) => hostMatches(host, s))) return false;
  return IN_APP_HOST_SUFFIXES.some((s) => hostMatches(host, s));
}

// Schemes handed to the OS. Anything else (javascript:, file:, data:, custom
// app schemes) is dropped rather than passed to xdg-open.
const EXTERNAL_SCHEMES = new Set(["http:", "https:", "mailto:", "tel:", "msteams:"]);

/**
 * Classifies a window.open / target=_blank request.
 * - "child":    open as a child window inside the app (Outlook pop-outs)
 * - "external": hand to the system browser or handler
 * - "deny":     drop it
 * @param {string} url
 * @returns {"child"|"external"|"deny"}
 */
function classifyWindowOpen(url) {
  // Outlook pop-out windows (compose, reading pane, calendar event) open
  // about:blank first and then navigate or write into it.
  if (url === "about:blank") return "child";
  if (isInAppUrl(url)) return "child";
  const parsed = parseUrl(url);
  if (parsed && EXTERNAL_SCHEMES.has(parsed.protocol)) return "external";
  return "deny";
}

/**
 * Classifies a same-window (main frame) navigation.
 * Federated sign-in sends the main frame to the organisation's own identity
 * provider (ADFS, Okta, Ping, ...) on hosts no list can predict, so any https
 * navigation stays in the app. Links in emails reach the user through
 * window.open, not here; Safe Links is the exception that can still land here.
 * @param {string} url
 * @returns {"allow"|"external"|"deny"}
 */
function classifyNavigation(url) {
  const parsed = parseUrl(url);
  if (!parsed) return "deny";
  if (parsed.protocol === "https:") {
    const host = parsed.hostname.toLowerCase();
    return EXTERNAL_HOST_SUFFIXES.some((s) => hostMatches(host, s)) ? "external" : "allow";
  }
  return EXTERNAL_SCHEMES.has(parsed.protocol) ? "external" : "deny";
}

module.exports = { isInAppUrl, classifyWindowOpen, classifyNavigation };
