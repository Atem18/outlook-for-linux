"use strict";

// Teams meeting links (the calendar's Join button, links in invites) open in
// a Teams desktop app such as teams-for-linux instead of the browser. The
// https link is rewritten to the msteams:// form Teams' own launcher page
// hands to the desktop app; teams-for-linux registers that scheme and turns
// it back into the https link it loads. Electron-free: the command runner is
// passed in, so it is unit tested. Never log the link: it identifies the
// meeting.

// Same hosts and paths as teams-for-linux's `msTeamsProtocols.v2`, so every
// link rewritten here is one it accepts.
const TEAMS_LINK =
  /^https:\/\/teams\.(?:microsoft\.com|live\.com|cloud\.microsoft)\/(?:meet\/|l\/(?:app|call|channel|chat|entity|file|meet(?:ing|up-join)|message|task|team)\/)/i;

const TEAMS_SCHEME_MIME = "x-scheme-handler/msteams";

/**
 * @param {string} url
 * @returns {string|null} the msteams:// form of a Teams link, or null
 */
function teamsAppUrl(url) {
  if (typeof url !== "string" || !TEAMS_LINK.test(url)) return null;
  return `msteams${url.slice("https".length)}`;
}

/**
 * Where a link goes, per the `teamsLinks` option:
 * - "teams":   always the msteams:// link
 * - "browser": never rewritten
 * - "auto":    the msteams:// link when the desktop has a handler for it.
 *              Inside Flatpak or Snap xdg-mime cannot see the host's
 *              handlers, so auto means the browser there.
 * @param {object} options
 * @param {string} options.url
 * @param {"auto"|"teams"|"browser"} options.mode
 * @param {boolean} options.sandboxed
 * @param {(command: string, args: string[]) => Promise<string>} options.run - runs a command, resolves its stdout
 * @returns {Promise<string>} the URL to hand to the OS
 */
async function resolveExternalUrl({ url, mode, sandboxed, run }) {
  const appUrl = teamsAppUrl(url);
  if (!appUrl || mode === "browser") return url;
  if (mode === "teams") return appUrl;
  if (sandboxed) return url;
  try {
    const handler = (await run("xdg-mime", ["query", "default", TEAMS_SCHEME_MIME])).trim();
    return handler ? appUrl : url;
  } catch {
    return url;
  }
}

module.exports = { TEAMS_SCHEME_MIME, teamsAppUrl, resolveExternalUrl };
