"use strict";

// Which URL to open at launch. Pure, so it is unit tested without Electron.
//
// The default URL sends everyone to Microsoft's converged sign-in, which is
// right the first time. On later launches it is wrong for personal accounts:
// login.microsoftonline.com does not pick up a personal session by itself, so
// a signed-in outlook.com user would be asked to sign in on every launch. The
// app therefore remembers which mailbox origin the user ended up in and
// starts there.

// Mailbox origins worth remembering. Personal accounts live on
// outlook.live.com; work and school accounts on the others.
const MAILBOX_ORIGINS = new Set([
  "https://outlook.live.com",
  "https://outlook.cloud.microsoft",
  "https://outlook.office.com",
  "https://outlook.office365.com",
]);

const MAILBOX_PATH = /^\/(mail|calendar|people|owa)(\/|$)/;

/**
 * The mailbox origin of a URL the main window has loaded, or null when it is
 * not a signed-in Outlook page (sign-in, marketing, consent pages).
 * @param {string} url
 * @returns {string|null}
 */
function mailboxOrigin(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return null;
  }
  if (!MAILBOX_ORIGINS.has(parsed.origin)) return null;
  return MAILBOX_PATH.test(parsed.pathname) ? parsed.origin : null;
}

/**
 * @param {string} configuredUrl - config.url
 * @param {string} defaultUrl - the schema default for url
 * @param {string|undefined} remembered - a previously stored mailbox origin
 * @returns {string}
 */
function startUrl(configuredUrl, defaultUrl, remembered) {
  // An explicitly configured URL always wins.
  if (configuredUrl !== defaultUrl) return configuredUrl;
  if (typeof remembered === "string" && MAILBOX_ORIGINS.has(remembered)) {
    return `${remembered}/mail/`;
  }
  return defaultUrl;
}

/**
 * Whether the main window has landed on a page that is a dead end after
 * signing out, from which the app should go back to sign-in:
 * - work and school sign-out ends on login.microsoftonline.com's "You signed
 *   out of your account" page (.../logout or .../logoutsession), which offers
 *   no way back (backported from teams-for-linux signOutReload.js),
 * - outlook.live.com sends signed-out visitors, including a remembered
 *   mailbox whose session expired, to Microsoft's marketing site,
 * - personal-account sign-out ends, after login.live.com, on www.msn.com.
 * @param {string} url
 * @returns {boolean}
 */
function isSignedOutPage(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:") return false;
  if (parsed.hostname === "login.microsoftonline.com") {
    return /\/logout(session)?$/.test(parsed.pathname);
  }
  return SIGNED_OUT_HOSTS.has(parsed.hostname);
}

const SIGNED_OUT_HOSTS = new Set(["www.microsoft.com", "www.msn.com"]);

module.exports = { mailboxOrigin, startUrl, isSignedOutPage };
