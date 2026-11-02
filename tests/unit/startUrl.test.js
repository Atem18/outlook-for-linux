const { test } = require("node:test");
const assert = require("node:assert/strict");
const { mailboxOrigin, startUrl, isSignedOutPage } = require("../../app/mainAppWindow/startUrl");

const DEFAULT = "https://outlook.cloud.microsoft/mail/";

test("mailboxOrigin recognises signed-in Outlook pages only", () => {
  assert.equal(mailboxOrigin("https://outlook.live.com/mail/0/inbox"), "https://outlook.live.com");
  assert.equal(mailboxOrigin("https://outlook.cloud.microsoft/calendar/view/week"), "https://outlook.cloud.microsoft");
  assert.equal(mailboxOrigin("https://outlook.live.com/"), null);
  assert.equal(mailboxOrigin("https://login.microsoftonline.com/common/oauth2/v2.0/authorize"), null);
  assert.equal(mailboxOrigin("https://outlook.live.com.evil.example/mail/"), null);
  assert.equal(mailboxOrigin("not a url"), null);
});

test("startUrl reopens the remembered mailbox", () => {
  assert.equal(startUrl(DEFAULT, DEFAULT, "https://outlook.live.com"), "https://outlook.live.com/mail/");
  assert.equal(startUrl(DEFAULT, DEFAULT, undefined), DEFAULT);
});

test("startUrl ignores an unknown remembered origin", () => {
  assert.equal(startUrl(DEFAULT, DEFAULT, "https://evil.example"), DEFAULT);
});

test("an explicitly configured url beats the remembered mailbox", () => {
  const custom = "https://outlook.office365.com/mail/";
  assert.equal(startUrl(custom, DEFAULT, "https://outlook.live.com"), custom);
});

// Cases shared with teams-for-linux tests/unit/signOutReload.test.js.
test("Microsoft's signed-out pages are dead ends to leave", () => {
  for (const url of [
    "https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri=x",
    "https://login.microsoftonline.com/common/oauth2/logoutsession",
    "https://www.microsoft.com/fr-fr/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook",
    "https://www.msn.com/fr-fr?ocid=mailsignout",
  ]) {
    assert.equal(isSignedOutPage(url), true, url);
  }
});

test("sign-in and mailbox pages are left alone", () => {
  for (const url of [
    "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    "https://login.microsoftonline.com/common/login",
    "https://outlook.live.com/mail/",
    "https://outlook.cloud.microsoft/owa/logout",
    "http://www.microsoft.com/",
    "about:blank",
  ]) {
    assert.equal(isSignedOutPage(url), false, url);
  }
});
