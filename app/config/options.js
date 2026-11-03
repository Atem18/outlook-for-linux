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
  defaultMailClient: {
    default: false,
    describe:
      "Make the app the default mail client at startup, so mailto: links open an Outlook compose window. Off by default to leave an existing mail client alone; inside Flatpak or Snap, choose the default in the desktop settings instead.",
    type: "boolean",
  },
  clientCertPath: {
    default: "",
    describe:
      "Client certificate file (PKCS#12, .p12/.pfx) to import at startup, for sign-ins that ask for a certificate. Linux only.",
    type: "string",
  },
  clientCertPassword: {
    default: "",
    describe: "Password of the clientCertPath file",
    type: "string",
  },
  clientCertPinDialog: {
    default: false,
    describe:
      "Show a PIN dialog when a smartcard (PKCS#11) client certificate needs unlocking. Without it such sign-ins fail silently. Linux only; the card must already be registered with NSS (e.g. OpenSC and modutil).",
    type: "boolean",
  },
  intune: {
    default: false,
    describe:
      "Single sign-on through the Microsoft Identity Broker on Intune-enrolled machines (work accounts): signs in with the device's credentials and satisfies device-compliance policies. Needs microsoft-identity-broker running.",
    type: "boolean",
  },
  intuneUser: {
    default: "",
    describe: "Broker account to use when it holds several (user@example.com); the first one when empty",
    type: "string",
  },
  webauthn: {
    default: false,
    describe:
      "Hardware security keys (FIDO2, e.g. YubiKey) for Microsoft sign-in, through fido2-tools (install the fido2-tools or libfido2 package). Linux only.",
    type: "boolean",
  },
  webauthnExtraOrigins: {
    default: [],
    describe:
      "Exact https origins, beyond Microsoft's sign-in pages, allowed to use hardware security keys, for organisations that sign in on their own identity provider, e.g. [\"https://sso.example.com\"].",
    type: "array",
  },
  webauthnDebug: {
    default: false,
    describe: "Verbose hardware security key logs, for troubleshooting",
    type: "boolean",
  },
  customCSSLocation: {
    default: "",
    describe:
      "Path to a CSS file injected into Outlook's pages, to restyle them. Empty for none.",
    type: "string",
  },
  globalShortcuts: {
    default: [],
    describe:
      "Shortcuts registered system-wide and forwarded to Outlook while the app is not focused, as Electron accelerators, e.g. [\"Control+Shift+M\"]. On Wayland, start with --enable-features=GlobalShortcutsPortal.",
    type: "array",
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
