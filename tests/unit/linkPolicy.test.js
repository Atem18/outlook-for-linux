const { test } = require("node:test");
const assert = require("node:assert/strict");
const {
  isInAppUrl,
  classifyWindowOpen,
  classifyNavigation,
} = require("../../app/mainAppWindow/linkPolicy");

test("Outlook and Microsoft sign-in hosts load in the app", () => {
  for (const url of [
    "https://outlook.cloud.microsoft/mail/",
    "https://outlook.office.com/calendar/view/week",
    "https://outlook.live.com/mail/0/inbox",
    "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    "https://login.live.com/oauth20_authorize.srf",
  ]) {
    assert.equal(isInAppUrl(url), true, url);
  }
});

test("lookalike and non-https hosts do not load in the app", () => {
  for (const url of [
    "https://outlook.office.com.evil.example/",
    "https://notoutlook.office.com/",
    "http://outlook.office.com/mail/",
    "https://example.com/",
    "not a url",
  ]) {
    assert.equal(isInAppUrl(url), false, url);
  }
});

test("Safe Links wrappers open externally even though they are outlook.com subdomains", () => {
  const url = "https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com";
  assert.equal(isInAppUrl(url), false);
  assert.equal(classifyWindowOpen(url), "external");
  assert.equal(classifyNavigation(url), "external");
});

test("window.open: pop-outs stay in the app, other links go external, unsafe schemes are dropped", () => {
  assert.equal(classifyWindowOpen("about:blank"), "child");
  assert.equal(classifyWindowOpen("https://outlook.office.com/mail/deeplink/compose"), "child");
  assert.equal(classifyWindowOpen("https://example.com/"), "external");
  assert.equal(classifyWindowOpen("mailto:someone@example.com"), "external");
  assert.equal(classifyWindowOpen("javascript:alert(1)"), "deny");
  assert.equal(classifyWindowOpen("file:///etc/passwd"), "deny");
});

test("navigation: federated identity providers stay in the app", () => {
  assert.equal(classifyNavigation("https://adfs.corp.example/adfs/ls/"), "allow");
  assert.equal(classifyNavigation("https://example.okta.com/login"), "allow");
  assert.equal(classifyNavigation("http://insecure.example/"), "external");
  assert.equal(classifyNavigation("file:///etc/passwd"), "deny");
});
