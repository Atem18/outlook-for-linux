const { test } = require("node:test");
const assert = require("node:assert/strict");
const { parseMailto, findMailtoArg, composeUrl } = require("../../app/mainAppWindow/mailtoLink");

const WORK = "https://outlook.cloud.microsoft";

test("parseMailto reads recipients, subject and body", () => {
  assert.deepEqual(
    parseMailto("mailto:a@example.com,b@example.com?cc=c@example.com&bcc=d@example.com&subject=Hello%20there&body=Line%201%0D%0ALine%202"),
    {
      to: ["a@example.com", "b@example.com"],
      cc: ["c@example.com"],
      bcc: ["d@example.com"],
      subject: "Hello there",
      body: "Line 1\r\nLine 2",
    },
  );
});

test("parseMailto keeps + literal, as RFC 6068 says", () => {
  const fields = parseMailto("mailto:user+tag@example.com?subject=1+1");
  assert.deepEqual(fields.to, ["user+tag@example.com"]);
  assert.equal(fields.subject, "1+1");
});

test("parseMailto: case-insensitive scheme and headers, to= adds up, unknown headers dropped", () => {
  const fields = parseMailto("MAILTO:?To=a@example.com&to=b@example.com&SUBJECT=x&subject=y&in-reply-to=%3Cid%3E");
  assert.deepEqual(fields.to, ["a@example.com", "b@example.com"]);
  assert.equal(fields.subject, "x");
  assert.equal(Object.keys(fields).includes("in-reply-to"), false);
});

test("parseMailto survives bad percent-encoding and empty links", () => {
  assert.equal(parseMailto("mailto:?subject=100%").subject, "100%");
  assert.deepEqual(parseMailto("mailto:"), { to: [], cc: [], bcc: [], subject: "", body: "" });
});

test("parseMailto ignores anything that is not a mailto: link", () => {
  assert.equal(parseMailto("https://example.com/?to=a"), null);
  assert.equal(parseMailto("--minimized"), null);
  assert.equal(parseMailto(undefined), null);
});

test("findMailtoArg picks the link out of an Electron command line", () => {
  const args = ["/opt/Outlook for Linux/outlook-for-linux", "--no-sandbox", "mailto:a@example.com"];
  assert.equal(findMailtoArg(args), "mailto:a@example.com");
  assert.equal(findMailtoArg(["/usr/bin/outlook-for-linux", "--minimized"]), null);
  assert.equal(findMailtoArg(undefined), null);
});

test("composeUrl builds Outlook's compose deeplink on the mailbox origin", () => {
  const url = new URL(composeUrl(parseMailto("mailto:a@example.com,user+tag@example.com?subject=Hi%20%26%20bye&body=x"), WORK));
  assert.equal(url.origin, WORK);
  assert.equal(url.pathname, "/mail/deeplink/compose");
  assert.equal(url.searchParams.get("to"), "a@example.com,user+tag@example.com");
  assert.equal(url.searchParams.get("subject"), "Hi & bye");
  assert.equal(url.searchParams.get("body"), "x");
  assert.equal(url.searchParams.has("cc"), false);
});

test("composeUrl encodes spaces as %20, since Outlook shows a + literally", () => {
  const url = composeUrl(parseMailto("mailto:a@example.com?subject=Hello%20world%20!&body=Lorem%20ipsum"), WORK);
  assert.match(url, /subject=Hello%20world%20!/);
  assert.match(url, /body=Lorem%20ipsum/);
  assert.doesNotMatch(url, /\+/);
});

test("composeUrl cannot be steered to another origin or path", () => {
  const url = new URL(composeUrl(parseMailto("mailto:a@example.com?subject=//evil.example/x"), "https://outlook.live.com"));
  assert.equal(url.origin, "https://outlook.live.com");
  assert.equal(url.pathname, "/mail/deeplink/compose");
});
