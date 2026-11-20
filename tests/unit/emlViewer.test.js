const { test } = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const {
  findEmlArgs,
  parseEml,
  toViewModel,
  messageDocument,
  replyFields,
  htmlToText,
  attachmentFileName,
} = require("../../app/emlViewer/message");

const CRLF = (lines) => lines.join("\r\n");

const PLAIN = CRLF([
  "From: =?UTF-8?Q?Ren=C3=A9e_Doe?= <renee@example.com>",
  "To: Bob <bob@example.com>, carol@example.com",
  "Cc: Team: dave@example.com, erin@example.com;",
  "Subject: =?UTF-8?B?SGVsbG8g8J+Riw==?=",
  "Date: Tue, 17 Nov 2026 09:30:00 +0100",
  "Content-Type: text/plain; charset=utf-8",
  "",
  "Hi Bob,",
  "See https://example.com/page. <b>not bold</b>",
  "",
]);

const RICH = CRLF([
  "From: alice@example.com",
  "Reply-To: replies@example.com",
  "To: bob@example.com",
  "Subject: Re: Report",
  "MIME-Version: 1.0",
  'Content-Type: multipart/mixed; boundary="outer"',
  "",
  "--outer",
  'Content-Type: multipart/related; boundary="inner"',
  "",
  "--inner",
  "Content-Type: text/html; charset=utf-8",
  "",
  '<p>Logo: <img src="cid:logo@x"></p><img src="https://tracker.example/p.gif">',
  "--inner",
  "Content-Type: image/png",
  "Content-ID: <logo@x>",
  "Content-Transfer-Encoding: base64",
  "",
  "iVBORw0KGgo=",
  "--inner--",
  "--outer",
  'Content-Type: application/pdf; name="../../etc/report.pdf"',
  "Content-Disposition: attachment; filename=\"../../etc/report.pdf\"",
  "Content-Transfer-Encoding: base64",
  "",
  "JVBERi0xLjQK",
  "--outer--",
  "",
]);

test("findEmlArgs picks .eml paths and file:// URIs, resolved against the given directory", () => {
  const cwd = "/home/user/Mail";
  assert.deepEqual(
    findEmlArgs(
      [
        "/opt/outlook-for-linux/outlook-for-linux",
        ".",
        "--trace-warnings",
        "inbox/one.eml",
        "/tmp/Two.EML",
        "file:///tmp/with%20space.eml",
        "mailto:a@example.com",
        "https://example.com/x.eml",
        "notes.txt",
        "inbox/one.eml",
      ],
      cwd,
    ),
    [path.join(cwd, "inbox/one.eml"), "/tmp/Two.EML", "/tmp/with space.eml"],
  );
  assert.deepEqual(findEmlArgs(undefined, cwd), []);
  assert.deepEqual(findEmlArgs(["--file=x.eml"], cwd), []);
});

test("a plain-text message: decoded headers, flattened groups, escaped and linked body", async () => {
  const view = toViewModel(await parseEml(Buffer.from(PLAIN)));
  assert.equal(view.subject, "Hello 👋");
  assert.deepEqual(view.from, ["Renée Doe <renee@example.com>"]);
  assert.deepEqual(view.to, ["Bob <bob@example.com>", "carol@example.com"]);
  assert.deepEqual(view.cc, ["dave@example.com", "erin@example.com"]);
  assert.equal(view.date, "2026-11-17T08:30:00.000Z");
  assert.match(view.html, /&lt;b&gt;not bold&lt;\/b&gt;/);
  assert.match(view.html, /<a href="https:\/\/example\.com\/page">https:\/\/example\.com\/page<\/a>\./);
  assert.equal(view.hasRemoteContent, false);
  assert.deepEqual(view.attachments, []);
});

test("an HTML message: cid images inlined, real attachments listed with safe names", async () => {
  const view = toViewModel(await parseEml(Buffer.from(RICH)));
  assert.match(view.html, /<img src="data:image\/png;base64,iVBORw0KGgo=">/);
  assert.doesNotMatch(view.html, /cid:/);
  assert.equal(view.hasRemoteContent, true);
  assert.deepEqual(
    view.attachments.map(({ filename, mimeType, size }) => ({ filename, mimeType, size })),
    [{ filename: "report.pdf", mimeType: "application/pdf", size: 9 }],
  );
});

test("attachmentFileName strips directories and control characters, with fallbacks", () => {
  assert.equal(attachmentFileName({ filename: "..\\..\\evil.exe", mimeType: "x" }, 0), "evil.exe");
  assert.equal(attachmentFileName({ filename: "a\nb.txt", mimeType: "text/plain" }, 0), "ab.txt");
  assert.equal(attachmentFileName({ filename: "..", mimeType: "x" }, 2), "attachment-3");
  assert.equal(attachmentFileName({ filename: null, mimeType: "text/calendar" }, 0), "attachment-1.ics");
  assert.equal(attachmentFileName({ filename: null, mimeType: "message/delivery-status" }, 1), "attachment-2.txt");
  assert.equal(attachmentFileName({ filename: null, mimeType: "text/x-anything" }, 0), "attachment-1.txt");
  assert.equal(attachmentFileName({ filename: null, mimeType: "application/octet-stream" }, 0), "attachment-1");
  assert.equal(attachmentFileName({ filename: null, mimeType: "message/rfc822" }, 0), "message-1.eml");
});

test("messageDocument blocks remote content unless asked, and sends links to a new window", () => {
  const blocked = messageDocument("<p>x</p>");
  assert.match(blocked, /default-src 'none'/);
  assert.match(blocked, /img-src data:;/);
  assert.doesNotMatch(blocked, /https:/);
  assert.match(blocked, /<base target="_blank">/);
  assert.ok(blocked.indexOf("Content-Security-Policy") < blocked.indexOf("<p>x</p>"));
  assert.match(messageDocument("", { remoteContent: true }), /img-src data: https: http:;/);
});

test("replyFields answers the Reply-To, adds Re: once and quotes the text", async () => {
  const plain = replyFields(await parseEml(Buffer.from(PLAIN)), "17 Nov 2026");
  assert.deepEqual(plain.to, ["renee@example.com"]);
  assert.equal(plain.subject, "Re: Hello 👋");
  assert.match(plain.body, /^\n\nOn 17 Nov 2026, Renée Doe <renee@example.com> wrote:\n> Hi Bob,\n/);

  const rich = replyFields(await parseEml(Buffer.from(RICH)));
  assert.deepEqual(rich.to, ["replies@example.com"]);
  assert.equal(rich.subject, "Re: Report");
  // HTML only: the quote comes from the HTML.
  assert.match(rich.body, /\n> Logo:$/);
});

test("htmlToText keeps the text and line breaks, drops styles and tags", () => {
  assert.equal(
    htmlToText("<html><head><style>p{}</style></head><body><p>One &amp; two</p><div>Three<br>4&#x2764;&nbsp;5</div><b>&bogus;</b></body></html>"),
    "One & two\nThree\n4❤ 5\n&bogus;",
  );
});
