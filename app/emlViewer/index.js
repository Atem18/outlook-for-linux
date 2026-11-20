const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const fs = require("node:fs/promises");
const path = require("node:path");
const {
  MAX_EML_BYTES,
  parseEml,
  toViewModel,
  messageDocument,
  replyFields,
  toBuffer,
} = require("./message");

// A session of its own, in memory only: remote content the user lets a
// message load never carries Outlook's cookies, and nothing stays on disk.
const VIEWER_PARTITION = "eml-viewer";

// Schemes a link in a message may open, in the browser or as a compose window.
const EXTERNAL_SCHEMES = new Set(["http:", "https:"]);

/**
 * Opens .eml files in a read-only viewer window, as Outlook for Windows does.
 * The message is parsed in the main process; the window shows its body in a
 * sandboxed iframe that runs no scripts and, unless the user asks, loads
 * nothing from the network.
 */
class EmlViewer {
  // webContents of each viewer window → its parsed message
  #messages = new Map();
  #openCompose;
  #openExternal;
  #icon;

  /**
   * @param {object} options
   * @param {(fields: object) => void} options.openCompose - opens a compose window for parsed mailto: fields
   * @param {(url: string) => void} options.openExternal - opens a link in the browser
   * @param {(url: string) => void} options.openMailto - opens a compose window for a mailto: link
   * @param {() => Electron.NativeImage|string} options.icon - the window icon
   */
  constructor({ openCompose, openExternal, openMailto, icon }) {
    this.#openCompose = openCompose;
    this.#openExternal = (url) => {
      if (/^mailto:/i.test(url)) openMailto(url);
      else openExternal(url);
    };
    this.#icon = icon;
    this.#registerIpcHandlers();
  }

  /**
   * Opens one file in a new viewer window. Errors are shown to the user,
   * never thrown.
   * @param {string} file - absolute path
   */
  async open(file) {
    let email;
    try {
      const stat = await fs.stat(file);
      if (!stat.isFile()) throw new Error("not a file");
      if (stat.size > MAX_EML_BYTES) throw new Error("file too large");
      email = await parseEml(await fs.readFile(file));
    } catch (err) {
      console.error("[EML] Could not open message", { error: err.code ?? err.message });
      dialog.showErrorBox("Could not open message", `${path.basename(file)} could not be read as an email message.`);
      return;
    }
    const view = toViewModel(email);
    console.info("[EML] Opening message", {
      attachments: view.attachments.length,
      hasRemoteContent: view.hasRemoteContent,
    });
    this.#createWindow(email, view);
  }

  /** File → Open Message…: picks .eml files and opens each. */
  async openDialog(parent) {
    const { canceled, filePaths } = await dialog.showOpenDialog(parent, {
      title: "Open Message",
      properties: ["openFile", "multiSelections"],
      filters: [
        { name: "Email messages", extensions: ["eml"] },
        { name: "All files", extensions: ["*"] },
      ],
    });
    if (canceled) return;
    for (const file of filePaths) await this.open(file);
  }

  #createWindow(email, view) {
    const win = new BrowserWindow({
      title: view.subject || "(no subject)",
      width: 900,
      height: 760,
      icon: this.#icon(),
      autoHideMenuBar: true,
      webPreferences: {
        partition: VIEWER_PARTITION,
        preload: path.join(__dirname, "viewerPreload.js"),
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
        webviewTag: false,
        spellcheck: false,
      },
    });
    const contents = win.webContents;
    this.#messages.set(contents, { email, view });
    contents.once("destroyed", () => this.#messages.delete(contents));

    // The viewer page itself never navigates, and neither does the
    // message's iframe (a <meta http-equiv="refresh"> or a link with its
    // own target). Links open through window.open (the iframe's <base
    // target="_blank">) and go to the browser.
    contents.on("will-navigate", (event) => event.preventDefault());
    contents.on("will-frame-navigate", (details) => {
      if (!details.isMainFrame && details.url !== "about:srcdoc") details.preventDefault();
    });
    contents.setWindowOpenHandler(({ url }) => {
      let scheme = null;
      try {
        scheme = new URL(url).protocol;
      } catch {
        // dropped below
      }
      if (EXTERNAL_SCHEMES.has(scheme) || scheme === "mailto:") this.#openExternal(url);
      else console.info("[EML] Link not opened", { scheme });
      return { action: "deny" };
    });
    // The page's title is fixed; keep the window named after the message.
    win.on("page-title-updated", (event) => event.preventDefault());
    win.loadFile(path.join(__dirname, "viewer.html")).catch(() => console.warn("[EML] Viewer failed to load"));
  }

  #entry(event) {
    return this.#messages.get(event.sender) ?? null;
  }

  #registerIpcHandlers() {
    ipcMain.handle("eml-viewer:message", (event) => {
      const entry = this.#entry(event);
      if (!entry) return null;
      const { html, ...header } = entry.view;
      return { ...header, document: messageDocument(html) };
    });

    ipcMain.handle("eml-viewer:remote-content", (event) => {
      const entry = this.#entry(event);
      if (!entry) return null;
      console.info("[EML] Loading remote content on request");
      return messageDocument(entry.view.html, { remoteContent: true });
    });

    ipcMain.handle("eml-viewer:save-attachment", async (event, index) => {
      const entry = this.#entry(event);
      const listed = entry?.view.attachments.find((attachment) => attachment.index === index);
      if (!listed) return false;
      const win = BrowserWindow.fromWebContents(event.sender);
      const { canceled, filePath } = await dialog.showSaveDialog(win, {
        title: "Save Attachment",
        defaultPath: path.join(app.getPath("downloads"), listed.filename),
      });
      if (canceled || !filePath) return false;
      try {
        await fs.writeFile(filePath, toBuffer(entry.email.attachments[index].content));
        console.info("[EML] Attachment saved", { bytes: listed.size });
        return true;
      } catch (err) {
        console.error("[EML] Could not save attachment", { error: err.code ?? err.message });
        dialog.showErrorBox("Could not save attachment", err.message);
        return false;
      }
    });

    ipcMain.handle("eml-viewer:reply", (event, dateText) => {
      const entry = this.#entry(event);
      if (!entry) return;
      this.#openCompose(replyFields(entry.email, typeof dateText === "string" ? dateText.slice(0, 200) : ""));
    });
  }
}

module.exports = EmlViewer;
