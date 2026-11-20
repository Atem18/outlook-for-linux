const { test } = require("node:test");
const assert = require("node:assert/strict");
const { teamsAppUrl, resolveExternalUrl } = require("../../app/mainAppWindow/teamsLink");

const JOIN = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%7d";

test("Teams meeting and deep links become msteams:// links", () => {
  assert.equal(teamsAppUrl(JOIN), JOIN.replace(/^https/, "msteams"));
  for (const url of [
    "https://teams.live.com/meet/9876543210?p=abc",
    "https://teams.microsoft.com/meet/123456789?p=xyz",
    "https://teams.cloud.microsoft/l/meetup-join/19%3ameeting_abc%40thread.v2/0",
    "https://teams.microsoft.com/l/chat/0/0?users=someone",
  ]) {
    assert.equal(teamsAppUrl(url), `msteams${url.slice(5)}`, url);
  }
});

test("other links are not rewritten", () => {
  for (const url of [
    "https://teams.microsoft.com/",
    "https://teams.microsoft.com/v2/?meetingjoin=true",
    "https://teams.microsoft.com.evil.example/l/meetup-join/x",
    "http://teams.microsoft.com/l/meetup-join/x",
    "https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fmeetup-join%2Fx",
    "https://example.com/l/meetup-join/x",
    "mailto:someone@example.com",
    "not a url",
  ]) {
    assert.equal(teamsAppUrl(url), null, url);
  }
  assert.equal(teamsAppUrl(undefined), null);
});

const handler = (stdout) => async () => stdout;
const failing = async () => {
  throw new Error("xdg-mime missing");
};

test("auto opens Teams links in the app only when a handler is registered", async () => {
  const app = JOIN.replace(/^https/, "msteams");
  const base = { url: JOIN, mode: "auto", sandboxed: false };
  assert.equal(await resolveExternalUrl({ ...base, run: handler("teams-for-linux.desktop\n") }), app);
  assert.equal(await resolveExternalUrl({ ...base, run: handler("\n") }), JOIN);
  assert.equal(await resolveExternalUrl({ ...base, run: failing }), JOIN);
  assert.equal(
    await resolveExternalUrl({ ...base, sandboxed: true, run: handler("teams-for-linux.desktop") }),
    JOIN,
  );
});

test("teams and browser modes ignore the handler lookup", async () => {
  let ran = false;
  const run = async () => {
    ran = true;
    return "";
  };
  const app = JOIN.replace(/^https/, "msteams");
  assert.equal(await resolveExternalUrl({ url: JOIN, mode: "teams", sandboxed: true, run }), app);
  assert.equal(await resolveExternalUrl({ url: JOIN, mode: "browser", sandboxed: false, run }), JOIN);
  assert.equal(ran, false);
});

test("non-Teams links pass through untouched", async () => {
  const url = "https://example.com/page";
  assert.equal(await resolveExternalUrl({ url, mode: "teams", sandboxed: false, run: failing }), url);
});
