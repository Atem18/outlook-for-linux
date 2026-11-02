import { test, expect, _electron as electron } from "@playwright/test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

test("launches and heads to Outlook or Microsoft sign-in", async ({}, testInfo) => {
  const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-e2e-"));
  const app = await electron.launch({
    args: [".", "--no-sandbox"],
    env: { ...process.env, E2E_USER_DATA_DIR: userDataDir },
  });

  try {
    const window = await app.firstWindow();
    // A clean profile has no session, so Outlook must redirect to Microsoft
    // sign-in; matching the start URL alone would pass before anything loads.
    await expect
      .poll(() => window.url(), { timeout: 30_000 })
      .toMatch(/^https:\/\/login\.(microsoftonline|live|microsoft)\.com\//);
    await window.screenshot({ path: testInfo.outputPath("launch.png") });
  } finally {
    await app.close();
    fs.rmSync(userDataDir, { recursive: true, force: true });
  }
});
