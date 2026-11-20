const { test } = require("node:test");
const assert = require("node:assert/strict");
const { applicationDirs, claimFileTypes, FILE_MIME_TYPES } = require("../../app/mainAppWindow/fileTypeDefault");

const ID = "outlook-for-linux.desktop";
const ENV = { HOME: "/home/u", XDG_DATA_DIRS: "/usr/local/share:/usr/share" };
const INSTALLED = (file) => file === "/usr/share/applications/outlook-for-linux.desktop";

// A fake xdg-mime: records calls, answers "query default" from `defaults`.
function fakeRun(defaults, { fail = false } = {}) {
  const calls = [];
  const run = async (command, args) => {
    calls.push([command, ...args]);
    if (fail) throw new Error("ENOENT");
    if (args[0] === "query") return `${defaults[args[2]] ?? ""}\n`;
    defaults[args[2]] = args[1];
    return "";
  };
  return { run, calls };
}

test("applicationDirs follows the XDG Base Directory defaults", () => {
  assert.deepEqual(applicationDirs({ HOME: "/home/u" }), [
    "/home/u/.local/share/applications",
    "/usr/local/share/applications",
    "/usr/share/applications",
  ]);
  assert.deepEqual(applicationDirs({ XDG_DATA_HOME: "/d", XDG_DATA_DIRS: "/a:" }), ["/d/applications", "/a/applications"]);
});

test("claims .eml when another app is the default", async () => {
  const { run, calls } = fakeRun({ "message/rfc822": "firefox_firefox.desktop" });
  const results = await claimFileTypes({ desktopId: ID, run, exists: INSTALLED, env: ENV });
  assert.deepEqual(FILE_MIME_TYPES, ["message/rfc822"]);
  assert.deepEqual(results, [{ mimeType: "message/rfc822", result: "claimed" }]);
  assert.deepEqual(calls, [
    ["xdg-mime", "query", "default", "message/rfc822"],
    ["xdg-mime", "default", ID, "message/rfc822"],
  ]);
});

test("leaves an existing default alone", async () => {
  const { run, calls } = fakeRun({ "message/rfc822": ID });
  const results = await claimFileTypes({ desktopId: ID, run, exists: INSTALLED, env: ENV });
  assert.deepEqual(results, [{ mimeType: "message/rfc822", result: "already" }]);
  assert.equal(calls.length, 1);
});

test("does nothing without an installed desktop entry (a run from a checkout)", async () => {
  const { run, calls } = fakeRun({});
  const results = await claimFileTypes({ desktopId: ID, run, exists: () => false, env: ENV });
  assert.deepEqual(results, [{ mimeType: "message/rfc822", result: "no-desktop-entry" }]);
  assert.equal(calls.length, 0);
});

test("reports a missing xdg-mime instead of throwing", async () => {
  const { run } = fakeRun({}, { fail: true });
  const results = await claimFileTypes({ desktopId: ID, run, exists: INSTALLED, env: ENV });
  assert.deepEqual(results, [{ mimeType: "message/rfc822", result: "failed" }]);
});
