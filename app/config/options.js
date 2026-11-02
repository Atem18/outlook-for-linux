// Configuration option definitions for Outlook for Linux.
//
// Single source of truth for the config schema, consumed by the yargs parser
// in ./index.js. Each entry carries a `default`, a `describe` string and a
// `type`. Keep it a plain data module with no Electron imports so unit tests
// and future doc generators can require it outside Electron.

module.exports = {
  // Like teams-for-linux, one URL serves every account type: it redirects to
  // Microsoft's converged sign-in (login.microsoftonline.com/common), which
  // accepts work, school and personal accounts and routes each to its mailbox.
  // outlook.live.com is not a usable default: signed out, it shows a
  // marketing page instead of a sign-in form.
  url: {
    default: "https://outlook.cloud.microsoft/mail/",
    describe: "Outlook web URL to load",
    type: "string",
  },
  partition: {
    default: "persist:outlook",
    describe:
      "Electron session partition holding cookies and storage. Use distinct values to keep separate instances signed in to different accounts.",
    type: "string",
  },
  chromeUserAgent: {
    default: `Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${process.versions.chrome} Safari/537.36`,
    describe:
      "User agent sent to Outlook. Electron's own user agent makes Outlook show an unsupported browser warning.",
    type: "string",
  },
  minimized: {
    default: false,
    describe: "Start the application minimized",
    type: "boolean",
  },
  titleBar: {
    default: "system",
    describe:
      "Title bar style. \"system\" uses the desktop's title bar; \"integrated\" hides it and draws the window controls over the Outlook top bar (client-side decorations). Linux and Windows; requires restart.",
    type: "string",
    choices: ["system", "integrated"],
  },
  trayIcon: {
    default: true,
    describe: "Show an icon in the system tray, with the unread count",
    type: "boolean",
  },
  closeToTray: {
    default: true,
    describe:
      "Closing the window hides it to the tray instead of quitting, so new-mail notifications and calendar reminders keep arriving. Ignored without trayIcon.",
    type: "boolean",
  },
  notifications: {
    default: true,
    describe:
      "Show Outlook's notifications (new mail, reminders) as native desktop notifications. Outlook's own notification settings still decide which ones it sends.",
    type: "boolean",
  },
  menubarAutoHide: {
    default: true,
    describe: "Hide the menu bar until Alt is pressed",
    type: "boolean",
  },
  spellCheckerLanguages: {
    default: [],
    describe:
      "Spell checker language codes, e.g. [\"en-US\", \"fr\"]. Empty uses the system locale.",
    type: "array",
  },
};
