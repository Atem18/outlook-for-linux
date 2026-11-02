const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { loadConfigFiles } = require("../../app/config");
const options = require("../../app/config/options");

function tempDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), "ofl-config-"));
}

test("the default URL is the account-neutral Outlook entry point", () => {
  assert.equal(options.url.default, "https://outlook.cloud.microsoft/mail/");
});

test("user config overrides system config key by key", () => {
  const userDir = tempDir();
  const systemFile = path.join(tempDir(), "config.json");
  fs.writeFileSync(systemFile, JSON.stringify({ partition: "persist:corp", minimized: true }));
  fs.writeFileSync(path.join(userDir, "config.json"), JSON.stringify({ partition: "persist:me" }));

  const { configFile, error } = loadConfigFiles(userDir, systemFile);
  assert.equal(error, null);
  assert.deepEqual(configFile, { partition: "persist:me", minimized: true });
});

test("an invalid user config is reported and ignored", () => {
  const userDir = tempDir();
  fs.writeFileSync(path.join(userDir, "config.json"), "{ not json");

  const { configFile, error } = loadConfigFiles(userDir, path.join(userDir, "missing.json"));
  assert.deepEqual(configFile, {});
  assert.ok(error);
});
