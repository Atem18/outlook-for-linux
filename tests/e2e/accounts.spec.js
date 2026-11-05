import { test, expect, _electron as electron } from "@playwright/test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

// multiAccount: a second account as a view over the window, on its own
// session, with the switcher in the bottom-left corner.
test("multiple accounts: one view per extra account, switched by Ctrl+Alt+<n>", async () => {
  const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), "ofl-e2e-accounts-"));
  fs.writeFileSync(
    path.join(userDataDir, "settings.json"),
    JSON.stringify({
      accounts: {
        list: [
          { id: "main", name: "Personal", partition: null },
          { id: "w1", name: "Work", partition: "persist:outlook-account-w1" },
        ],
        activeId: "w1",
      },
    })
  );
  const app = await electron.launch({
    args: [".", "--no-sandbox", "--multiAccount"],
    env: { ...process.env, E2E_USER_DATA_DIR: userDataDir },
  });

  const layout = () =>
    app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0];
      const [width, height] = win.getContentSize();
      return {
        width,
        height,
        mainPartition: win.webContents.session.getStoragePath()?.split("/").pop(),
        views: win.contentView.children.map((view) => ({
          bounds: view.getBounds(),
          visible: view.getVisible(),
          partition: view.webContents.session.getStoragePath()?.split("/").pop(),
          url: view.webContents.getURL(),
        })),
      };
    });

  try {
    await app.firstWindow();
    // The Work account's view, then the switcher on top.
    await expect.poll(async () => (await layout()).views.length, { timeout: 15_000 }).toBe(2);
    await expect
      .poll(async () => (await layout()).views[0].url, { timeout: 30_000 })
      .toMatch(/^https:\/\/login\.(microsoftonline|live|microsoft)\.com\//);

    const start = await layout();
    const [work, switcher] = start.views;
    expect(work.partition).toBe("outlook-account-w1");
    expect(start.mainPartition).toBe("outlook");
    expect(work.visible).toBe(true);
    expect(work.bounds).toEqual({ x: 0, y: 0, width: start.width, height: start.height });
    expect(switcher.url).toMatch(/accountSwitcher\/switcher\.html$/);
    expect(switcher.bounds).toEqual({ x: 0, y: start.height - 48, width: 48, height: 48 });

    // Ctrl+Alt+1 shows the main account (the window's own page) again.
    await app.evaluate(({ BrowserWindow }) => {
      const view = BrowserWindow.getAllWindows()[0].contentView.children[0];
      view.webContents.sendInputEvent({ type: "keyDown", keyCode: "1", modifiers: ["control", "alt"] });
    });
    await expect.poll(async () => (await layout()).views[0].visible).toBe(false);
    const settings = JSON.parse(fs.readFileSync(path.join(userDataDir, "settings.json"), "utf8"));
    expect(settings.accounts.activeId).toBe("main");
  } finally {
    await app.close();
    fs.rmSync(userDataDir, { recursive: true, force: true });
  }
});
