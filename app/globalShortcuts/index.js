/**
 * Global Shortcuts Module, ported from teams-for-linux app/globalShortcuts.
 *
 * Registers system-wide keyboard shortcuts that forward events to Outlook, so
 * Outlook's own shortcuts (e.g. Control+N for a new message) work while the
 * app is not focused. On Wayland, Electron registers them through the
 * desktop's GlobalShortcuts portal only when started with
 * --enable-features=GlobalShortcutsPortal.
 */

const { globalShortcut } = require("electron");
const os = require("node:os");

const isMac = os.platform() === "darwin";
let isRegistered = false;

/**
 * Parses an Electron accelerator string into key and modifiers.
 *
 * @param {string} accelerator - The accelerator string (e.g., "Control+Shift+M")
 * @returns {Object} - Object with key and modifiers array
 */
function parseAccelerator(accelerator) {
  const parts = accelerator.split("+");
  const modifiers = [];
  let key = "";

  for (const part of parts) {
    const lower = part.toLowerCase();
    if (lower === "commandorcontrol" || lower === "cmdorctrl") {
      modifiers.push(isMac ? "cmd" : "control");
    } else if (lower === "command" || lower === "cmd") {
      modifiers.push("cmd");
    } else if (lower === "control" || lower === "ctrl") {
      modifiers.push("control");
    } else if (lower === "shift") {
      modifiers.push("shift");
    } else if (lower === "alt" || lower === "option") {
      modifiers.push("alt");
    } else {
      // This is the key
      key = part;
    }
  }

  return { key, modifiers };
}

/**
 * @param {Electron.WebContents} contents - The page to send the event to
 * @param {string} accelerator - The accelerator string (e.g., "Control+Shift+M")
 */
function sendKeyboardEvent(contents, accelerator) {
  try {
    const parsed = parseAccelerator(accelerator);

    contents.sendInputEvent({
      type: "keyDown",
      keyCode: parsed.key,
      modifiers: parsed.modifiers
    });

    contents.sendInputEvent({
      type: "keyUp",
      keyCode: parsed.key,
      modifiers: parsed.modifiers
    });

    console.debug(`[GLOBAL_SHORTCUTS] Forwarded keyboard event: ${accelerator}`);
    return true;
  } catch (err) {
    console.error(`[GLOBAL_SHORTCUTS] Error sending keyboard event for ${accelerator}: ${err.message}`);
    return false;
  }
}

/**
 * @param {{globalShortcuts?: string[]}} config
 * @param {() => Electron.WebContents|null} getContents - the visible
 *   account's Outlook page
 * @param {Electron.App} app
 */
function register(config, getContents, app) {
  // Guard against multiple registrations
  if (isRegistered) {
    console.debug("[GLOBAL_SHORTCUTS] Already registered, skipping");
    return;
  }

  if (!Array.isArray(config.globalShortcuts) || config.globalShortcuts.length === 0) {
    console.debug("[GLOBAL_SHORTCUTS] No global shortcuts configured");
    isRegistered = true; // Mark as registered even with no shortcuts to maintain guard integrity
    return;
  }

  let registeredCount = 0;

  for (const shortcut of config.globalShortcuts) {
    // Skip empty or invalid shortcuts
    if (!shortcut || typeof shortcut !== "string") {
      console.debug(`[GLOBAL_SHORTCUTS] Skipping invalid shortcut: ${shortcut}`);
      continue;
    }

    try {
      const registered = globalShortcut.register(shortcut, () => {
        console.debug(`[GLOBAL_SHORTCUTS] Shortcut triggered: ${shortcut}`);

        const contents = getContents();
        if (contents && !contents.isDestroyed()) {
          // Forward the keyboard event to Outlook by simulating the key
          // press; Outlook handles it with its built-in keyboard shortcuts.
          sendKeyboardEvent(contents, shortcut);
        } else {
          console.warn(`[GLOBAL_SHORTCUTS] Main window not available for shortcut: ${shortcut}`);
        }
      });

      if (registered) {
        console.info(`[GLOBAL_SHORTCUTS] Registered: ${shortcut}`);
        registeredCount++;
      } else {
        console.warn(`[GLOBAL_SHORTCUTS] Failed to register ${shortcut} (may already be in use by another application)`);
      }
    } catch (err) {
      console.error(`[GLOBAL_SHORTCUTS] Error registering ${shortcut}: ${err.message}`);
    }
  }

  // Unregister all shortcuts on app quit
  app.on("will-quit", () => {
    globalShortcut.unregisterAll();
    console.debug("[GLOBAL_SHORTCUTS] Unregistered all shortcuts");
    isRegistered = false;
  });

  if (registeredCount > 0) {
    isRegistered = true;
    console.info(`[GLOBAL_SHORTCUTS] Successfully registered ${registeredCount} global shortcut(s)`);
  }
}

module.exports = { register, sendKeyboardEvent, parseAccelerator };
