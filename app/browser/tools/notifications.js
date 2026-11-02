/**
 * Native notifications, renderer half.
 *
 * Outlook creates web notifications (new mail, reminders) with
 * `new Notification(...)`. The preload runs in an isolated world, so it
 * replaces the page's `Notification` from the main world through
 * contextBridge: `installShim` is serialised and run there before Outlook's
 * own scripts, and talks back through a small bridge object. The main process
 * (app/notifications/index.js) shows the native notification and relays
 * clicks, so Outlook's own click handler opens the right message.
 */

const BRIDGE_KEY = "__oflNotifications";

// Runs in the page's main world: serialised by contextBridge, so it must be
// self-contained and can only reach the preload through the bridge.
function installShim(bridgeKey) {
  const bridge = globalThis[bridgeKey];
  if (!bridge) return;
  const MAX_LIVE = 50;
  const live = new Map();
  let counter = 0;

  bridge.onEvent((id, type) => {
    const notification = live.get(id);
    if (!notification) return;
    const event = new Event(type);
    notification.dispatchEvent(event);
    const handler = notification[`on${type}`];
    if (typeof handler === "function") {
      try {
        handler.call(notification, event);
      } catch (err) {
        console.error(err);
      }
    }
    if (type === "close" || type === "error") live.delete(id);
  });

  class Notification extends EventTarget {
    static get permission() {
      return "granted";
    }

    static requestPermission(callback) {
      if (typeof callback === "function") callback("granted");
      return Promise.resolve("granted");
    }

    static get maxActions() {
      return 0;
    }

    constructor(title, options = {}) {
      super();
      this.title = String(title ?? "");
      this.body = String(options.body ?? "");
      this.tag = String(options.tag ?? "");
      this.icon = String(options.icon ?? "");
      this.data = options.data ?? null;
      this.silent = options.silent === true;
      this.requireInteraction = options.requireInteraction === true;
      this.onclick = null;
      this.onshow = null;
      this.onclose = null;
      this.onerror = null;
      counter += 1;
      this.__id = `n${Date.now().toString(36)}${counter}`;
      live.set(this.__id, this);
      while (live.size > MAX_LIVE) live.delete(live.keys().next().value);
      bridge.show({
        id: this.__id,
        title: this.title,
        body: this.body,
        silent: this.silent,
        requireInteraction: this.requireInteraction,
      });
    }

    close() {
      bridge.close(this.__id);
    }
  }

  Object.defineProperty(globalThis, "Notification", {
    value: Notification,
    writable: true,
    configurable: true,
  });
}

/**
 * Called from the preload at load time, before the page's scripts.
 * @param {object} config - renderer config (`notifications` flag)
 * @param {Electron.IpcRenderer} ipcRenderer
 * @param {Electron.ContextBridge} contextBridge
 */
function init(config, ipcRenderer, contextBridge) {
  if (!config?.notifications || !contextBridge) return;

  const listeners = [];
  ipcRenderer.on("notification-event", (_event, payload) => {
    if (!payload || typeof payload.id !== "string" || typeof payload.type !== "string") return;
    for (const listener of listeners) listener(payload.id, payload.type);
  });

  // The page sees only these three functions; the main process validates
  // everything they send.
  contextBridge.exposeInMainWorld(BRIDGE_KEY, {
    show: (options) => ipcRenderer.send("notification-show", options),
    close: (id) => ipcRenderer.send("notification-close", id),
    onEvent: (listener) => {
      if (typeof listener === "function") listeners.push(listener);
    },
  });
  contextBridge.executeInMainWorld({ func: installShim, args: [BRIDGE_KEY] });
}

module.exports = { init, _internals: { installShim, BRIDGE_KEY } };
