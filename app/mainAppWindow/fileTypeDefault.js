"use strict";

// Makes the app the default for file types (.eml) with xdg-mime, as
// app.setAsDefaultProtocolClient does for mailto: with xdg-settings. Electron
// has no API for file types. Electron-free: the command runner and the file
// check are passed in, so it is unit tested.

const path = require("node:path");

// File types the desktop entries declare. Keep in step with
// `build.linux.mimeTypes` in package.json and MimeType in the Flatpak
// desktop entry.
const FILE_MIME_TYPES = ["message/rfc822"];

/**
 * Where the desktop looks for desktop entries (XDG Base Directory spec):
 * $XDG_DATA_HOME, then each of $XDG_DATA_DIRS, each with "applications".
 * @param {NodeJS.ProcessEnv} env
 * @returns {string[]}
 */
function applicationDirs(env) {
  const home = env.XDG_DATA_HOME || (env.HOME ? path.join(env.HOME, ".local", "share") : null);
  const system = (env.XDG_DATA_DIRS || "/usr/local/share:/usr/share").split(":").filter(Boolean);
  return [home, ...system].filter(Boolean).map((dir) => path.join(dir, "applications"));
}

/**
 * Claims each file type for the desktop entry, unless it already is the
 * default. Skipped when no installed desktop entry has that id (a dev run
 * from a checkout): a default naming a missing entry would only break the
 * user's file associations.
 * @param {object} options
 * @param {string} options.desktopId - e.g. "outlook-for-linux.desktop"
 * @param {string[]} [options.mimeTypes]
 * @param {(command: string, args: string[]) => Promise<string>} options.run - runs a command, resolves its stdout
 * @param {(file: string) => boolean} options.exists
 * @param {NodeJS.ProcessEnv} options.env
 * @returns {Promise<Array<{mimeType: string, result: "already"|"claimed"|"failed"|"no-desktop-entry"}>>}
 */
async function claimFileTypes({ desktopId, mimeTypes = FILE_MIME_TYPES, run, exists, env }) {
  const installed = applicationDirs(env).some((dir) => exists(path.join(dir, desktopId)));
  if (!installed) return mimeTypes.map((mimeType) => ({ mimeType, result: "no-desktop-entry" }));
  const results = [];
  for (const mimeType of mimeTypes) {
    try {
      const current = (await run("xdg-mime", ["query", "default", mimeType])).trim();
      if (current === desktopId) {
        results.push({ mimeType, result: "already" });
        continue;
      }
      await run("xdg-mime", ["default", desktopId, mimeType]);
      results.push({ mimeType, result: "claimed" });
    } catch {
      results.push({ mimeType, result: "failed" });
    }
  }
  return results;
}

module.exports = { FILE_MIME_TYPES, applicationDirs, claimFileTypes };
