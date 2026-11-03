"use strict";

/**
 * Owns a session's webRequest.onBeforeSendHeaders listener and shares it.
 * Electron keeps one such listener per session (registering again replaces
 * it), so app/unread (borrowing Outlook's bearer token) and app/intune
 * (adding the broker's SSO cookie to sign-in requests) would otherwise
 * silently knock each other out. Electron-free, so it is unit tested with a
 * fake session.
 */

const registries = new WeakMap();

class RequestHeaders {
  #session;
  #handlers = [];

  /** @param {Electron.Session} session */
  constructor(session) {
    this.#session = session;
  }

  /**
   * The registry for a session, created on first use.
   * @param {Electron.Session} session
   * @returns {RequestHeaders}
   */
  static for(session) {
    let registry = registries.get(session);
    if (!registry) {
      registry = new RequestHeaders(session);
      registries.set(session, registry);
    }
    return registry;
  }

  /**
   * Adds a handler. It runs for requests matching both its Electron URL
   * patterns (which narrow the listener's filter) and `matches`, and may
   * change details.requestHeaders in place, synchronously or not. Handlers
   * run one after another in the order they were added; a handler that throws
   * is skipped, never blocking the request.
   * @param {{urls: string[], matches: (url: string) => boolean, handle: (details: object) => (void|Promise<void>), name: string}} handler
   */
  use(handler) {
    this.#handlers.push(handler);
    const urls = [...new Set(this.#handlers.flatMap((h) => h.urls))];
    this.#session.webRequest.onBeforeSendHeaders({ urls }, (details, callback) => {
      this.#run(details).finally(() => callback({ requestHeaders: details.requestHeaders }));
    });
  }

  async #run(details) {
    for (const handler of this.#handlers) {
      if (!handler.matches(details.url)) continue;
      try {
        await handler.handle(details);
      } catch (err) {
        console.warn("[REQUEST_HEADERS] Handler failed", { handler: handler.name, error: err?.message });
      }
    }
  }
}

module.exports = RequestHeaders;
