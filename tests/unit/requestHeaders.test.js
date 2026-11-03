const { test } = require("node:test");
const assert = require("node:assert/strict");
const RequestHeaders = require("../../app/requestHeaders");

function fakeSession() {
  const session = { registrations: 0, filter: null, listener: null };
  session.webRequest = {
    onBeforeSendHeaders(filter, listener) {
      session.registrations += 1;
      session.filter = filter;
      session.listener = listener;
    },
  };
  return session;
}

function send(session, url) {
  return new Promise((resolve) => {
    session.listener({ url, requestHeaders: { Accept: "*/*" } }, (result) => resolve(result.requestHeaders));
  });
}

test("one listener serves every handler, with the union of their URLs", async () => {
  const session = fakeSession();
  const registry = RequestHeaders.for(session);
  assert.equal(RequestHeaders.for(session), registry);
  registry.use({ name: "a", urls: ["https://*/owa/service.svc*"], matches: (u) => u.includes("/owa/"), handle: (d) => { d.requestHeaders.A = "1"; } });
  registry.use({
    name: "b",
    urls: ["https://login.microsoftonline.com/*"],
    matches: (u) => u.startsWith("https://login."),
    handle: async (d) => { d.requestHeaders.B = "2"; },
  });
  assert.deepEqual(session.filter.urls, ["https://*/owa/service.svc*", "https://login.microsoftonline.com/*"]);

  assert.deepEqual(await send(session, "https://outlook.office.com/owa/service.svc?action=x"), { Accept: "*/*", A: "1" });
  assert.deepEqual(await send(session, "https://login.microsoftonline.com/common/oauth2"), { Accept: "*/*", B: "2" });
});

test("a failing handler does not block the request or the others", async (t) => {
  t.mock.method(console, "warn", () => {});
  const session = fakeSession();
  const registry = RequestHeaders.for(session);
  registry.use({ name: "broken", urls: ["https://*/*"], matches: () => true, handle: async () => { throw new Error("boom"); } });
  registry.use({ name: "ok", urls: ["https://*/*"], matches: () => true, handle: (d) => { d.requestHeaders.Ok = "yes"; } });
  assert.deepEqual(await send(session, "https://example.com/"), { Accept: "*/*", Ok: "yes" });
});
