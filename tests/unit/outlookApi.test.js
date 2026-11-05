const { test } = require("node:test");
const assert = require("node:assert/strict");
const {
  apiOriginOf,
  bearerFrom,
  getFolderUrl,
  getFolderBody,
  parseUnreadCount,
} = require("../../app/unread/outlookApi");

test("only Outlook's own service.svc calls lend their token", () => {
  assert.equal(apiOriginOf("https://outlook.cloud.microsoft/owa/service.svc?action=FindFolder"), "https://outlook.cloud.microsoft");
  assert.equal(apiOriginOf("https://outlook.live.com/owa/service.svc?action=GetFolder"), "https://outlook.live.com");
  assert.equal(apiOriginOf("https://outlook.cloud.microsoft/owa/notificationchannel"), null);
  assert.equal(apiOriginOf("https://evil.example/owa/service.svc"), null);
  assert.equal(apiOriginOf("https://outlook.cloud.microsoft.evil.example/owa/service.svc"), null);
  assert.equal(apiOriginOf("http://outlook.live.com/owa/service.svc"), null);
  assert.equal(apiOriginOf("not a url"), null);
});

test("reads a bearer Authorization header whatever its case", () => {
  assert.equal(bearerFrom({ Authorization: "Bearer abc.def" }), "Bearer abc.def");
  assert.equal(bearerFrom({ authorization: "Bearer abc" }), "Bearer abc");
  assert.equal(bearerFrom({ Authorization: "Basic abc" }), null);
  assert.equal(bearerFrom({ Authorization: "Bearer a b" }), null);
  assert.equal(bearerFrom(undefined), null);
});

test("asks GetFolder for the Inbox unread count only, on the token's origin", () => {
  assert.equal(getFolderUrl("https://outlook.live.com"), "https://outlook.live.com/owa/service.svc?action=GetFolder&app=Mail");
  const body = JSON.parse(getFolderBody());
  assert.equal(body.Body.FolderIds[0].Id, "inbox");
  assert.deepEqual(body.Body.FolderShape.AdditionalProperties.map((p) => p.FieldURI), ["folder:UnreadCount"]);
  assert.equal(body.Body.FolderShape.BaseShape, "IdOnly");
});

test("parses the unread count from a GetFolder response", () => {
  const response = { Body: { ResponseMessages: { Items: [{ Folders: [{ UnreadCount: 3 }] }] } } };
  assert.equal(parseUnreadCount(response), 3);
  assert.equal(parseUnreadCount({ Body: { ResponseMessages: { Items: [{ ResponseClass: "Error" }] } } }), null);
  assert.equal(parseUnreadCount(null), null);
});

test("only Outlook's own changes ask for the count again, not GetFolder", () => {
  const { isGetFolder } = require("../../app/unread/outlookApi");
  assert.equal(isGetFolder("https://outlook.live.com/owa/service.svc?action=GetFolder&app=Mail"), true);
  assert.equal(isGetFolder("https://outlook.live.com/owa/service.svc?action=getfolder"), true);
  assert.equal(isGetFolder("https://outlook.live.com/owa/service.svc?action=UpdateItem&app=Mail"), false);
  assert.equal(isGetFolder("https://outlook.live.com/owa/service.svc"), false);
  assert.equal(isGetFolder("not a url"), false);
});

test("the newest of the API and folder-pane counts wins", () => {
  const { latestCount } = require("../../app/unread/outlookApi");
  const fresh = 60_000;
  // Reading a message: the pane drops before the next poll.
  assert.equal(latestCount({ count: 5, at: 1000 }, { count: 4, at: 2000 }, 3000, fresh), 4);
  // The next poll answers after the pane.
  assert.equal(latestCount({ count: 4, at: 3000 }, { count: 5, at: 2000 }, 3500, fresh), 4);
  // No pane (narrow window): the API alone.
  assert.equal(latestCount({ count: 7, at: 1000 }, { count: null, at: 0 }, 2000, fresh), 7);
  // A stale API answer gives way to the pane, or to nothing.
  assert.equal(latestCount({ count: 7, at: 0 }, { count: 2, at: 0 }, 100_000, fresh), 2);
  assert.equal(latestCount({ count: 7, at: 0 }, { count: null, at: 0 }, 100_000, fresh), null);
  assert.equal(latestCount({ count: null, at: 0 }, { count: null, at: 0 }, 0, fresh), null);
});
