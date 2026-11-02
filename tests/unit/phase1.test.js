const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

const { trayTooltip, isBadgeDataUrl, sanitizeUnreadUpdate, compositeBadge } = require("../../app/tray/badge");
const { sanitizeNotification, isNotificationId } = require("../../app/notifications/sanitize");
const notificationsTool = require("../../app/browser/tools/notifications");
const unreadBadgeTool = require("../../app/browser/tools/unreadBadge");
const options = require("../../app/config/options");

describe("phase 1 defaults", () => {
  it("keeps Outlook running in the tray with notifications on", () => {
    assert.equal(options.trayIcon.default, true);
    assert.equal(options.closeToTray.default, true);
    assert.equal(options.notifications.default, true);
  });
});

describe("tray badge helpers", () => {
  it("puts the count in the tooltip only when there is unread mail", () => {
    assert.equal(trayTooltip(0), "Outlook for Linux");
    assert.equal(trayTooltip(7), "Outlook for Linux (7)");
  });

  it("accepts only reasonable PNG data URLs as badges", () => {
    assert.equal(isBadgeDataUrl("data:image/png;base64,iVBORw0KGgo="), true);
    assert.equal(isBadgeDataUrl("data:image/svg+xml;base64,PHN2Zz4="), false);
    assert.equal(isBadgeDataUrl("https://example.com/a.png"), false);
    assert.equal(isBadgeDataUrl(`data:image/png;base64,${"A".repeat(300_000)}`), false);
  });

  it("validates unread updates and drops a bad badge but keeps the count", () => {
    assert.deepEqual(sanitizeUnreadUpdate({ count: 3, badge: "data:image/png;base64,AAAA" }), {
      count: 3,
      badge: "data:image/png;base64,AAAA",
    });
    assert.deepEqual(sanitizeUnreadUpdate({ count: 3, badge: "javascript:x" }), { count: 3, badge: null });
    assert.equal(sanitizeUnreadUpdate({ count: -1 }), null);
    assert.equal(sanitizeUnreadUpdate({ count: 1.5 }), null);
    assert.equal(sanitizeUnreadUpdate(null), null);
  });

  it("composites the badge over the icon with premultiplied alpha", () => {
    const base = Buffer.from([100, 100, 100, 255, 100, 100, 100, 255, 0, 0, 0, 0]);
    // opaque red, transparent, half-transparent white (premultiplied)
    const badge = Buffer.from([0, 0, 255, 255, 0, 0, 0, 0, 128, 128, 128, 128]);
    assert.deepEqual([...compositeBadge(base, badge)], [0, 0, 255, 255, 100, 100, 100, 255, 128, 128, 128, 128]);
    assert.equal(compositeBadge(base, Buffer.alloc(4)), null);
  });
});

describe("notification validation", () => {
  it("keeps the fields Electron needs and caps their length", () => {
    const n = sanitizeNotification({ id: "n1", title: "T".repeat(1000), body: "B", silent: true, extra: 1 });
    assert.equal(n.title.length, 256);
    assert.deepEqual({ ...n, title: undefined }, { id: "n1", title: undefined, body: "B", silent: true, requireInteraction: false });
  });

  it("rejects a missing title or an unsafe id", () => {
    assert.equal(sanitizeNotification({ id: "n1", title: "" }), null);
    assert.equal(sanitizeNotification({ id: "../x", title: "T" }), null);
    assert.equal(isNotificationId("n1"), true);
    assert.equal(isNotificationId({}), false);
  });
});

// The shims run in the page's main world; here they run against fake globals
// and a fake bridge, the same shape the preload exposes.
function withGlobals(globals, fn) {
  const saved = {};
  for (const key of Object.keys(globals)) {
    saved[key] = Object.getOwnPropertyDescriptor(globalThis, key);
    Object.defineProperty(globalThis, key, { value: globals[key], configurable: true, writable: true });
  }
  try {
    return fn();
  } finally {
    for (const [key, descriptor] of Object.entries(saved)) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else delete globalThis[key];
    }
  }
}

describe("Notification shim", () => {
  const { installShim, BRIDGE_KEY } = notificationsTool._internals;

  function setup() {
    const shown = [];
    const closed = [];
    let emit = null;
    const bridge = {
      show: (o) => shown.push(o),
      close: (id) => closed.push(id),
      onEvent: (fn) => { emit = fn; },
    };
    return { bridge, shown, closed, emit: (...a) => emit(...a) };
  }

  it("forwards new notifications and reports granted permission", () => {
    const { bridge, shown, closed } = setup();
    withGlobals({ [BRIDGE_KEY]: bridge, Notification: undefined }, () => {
      installShim(BRIDGE_KEY);
      assert.equal(globalThis.Notification.permission, "granted");
      const n = new globalThis.Notification("New mail", { body: "Hello", tag: "x" });
      assert.equal(shown.length, 1);
      assert.equal(shown[0].title, "New mail");
      assert.equal(shown[0].body, "Hello");
      assert.match(shown[0].id, /^[A-Za-z0-9_-]+$/);
      n.close();
      assert.deepEqual(closed, [shown[0].id]);
    });
  });

  it("relays a click to the page's onclick and click listeners", async () => {
    const { bridge, shown, emit } = setup();
    await withGlobals({ [BRIDGE_KEY]: bridge, Notification: undefined }, async () => {
      installShim(BRIDGE_KEY);
      const n = new globalThis.Notification("Reminder");
      const seen = [];
      n.onclick = () => seen.push("onclick");
      n.addEventListener("click", () => seen.push("listener"));
      emit(shown[0].id, "click");
      assert.deepEqual(seen.sort(), ["listener", "onclick"]);
      assert.equal(await globalThis.Notification.requestPermission(), "granted");
    });
  });
});

describe("Inbox unread count", () => {
  const { readInboxUnread, badgeLabel, INBOX_SELECTOR } = unreadBadgeTool._internals;

  // Just enough of the DOM for readInboxUnread: leaf spans with text.
  const span = (text) => ({ children: [], textContent: text });
  const wrapper = { children: [{}], textContent: "Inbox3" };
  const pane = (spans) => ({
    querySelector: (selector) => (selector === INBOX_SELECTOR && spans ? { querySelectorAll: () => spans } : null),
  });

  it("reads the count shown next to the Inbox", () => {
    assert.equal(readInboxUnread(pane([wrapper, span("Inbox"), span(" 3 ")])), 3);
  });

  it("is 0 when the Inbox shows no count", () => {
    assert.equal(readInboxUnread(pane([span("Inbox")])), 0);
  });

  it("is unknown when the folder pane is not rendered", () => {
    assert.equal(readInboxUnread(pane(null)), null);
  });

  it("finds the Inbox by its language-independent folder name", () => {
    assert.match(INBOX_SELECTOR, /data-folder-name="inbox" i/);
  });

  it("caps the label at 99+", () => {
    assert.equal(badgeLabel(7), "7");
    assert.equal(badgeLabel(99), "99");
    assert.equal(badgeLabel(100), "99+");
  });
});

describe("connection retry policy", () => {
  const { isNetworkError, retryDelay, MAX_DELAY_MS } = require("../../app/connection/retryPolicy");

  it("retries only network failures", () => {
    assert.equal(isNetworkError(-106), true, "INTERNET_DISCONNECTED");
    assert.equal(isNetworkError(-105), true, "NAME_NOT_RESOLVED");
    assert.equal(isNetworkError(-200), false, "CERT_COMMON_NAME_INVALID is not retried");
    assert.equal(isNetworkError(-3), false, "ABORTED (a navigation replaced) is not retried");
  });

  it("backs off exponentially up to a minute", () => {
    assert.deepEqual([0, 1, 2, 3].map(retryDelay), [2000, 4000, 8000, 16000]);
    assert.equal(retryDelay(50), MAX_DELAY_MS);
  });
});
