const { WebContentsView, dialog, ipcMain, session } = require("electron");
const path = require("node:path");
const { showSecurePrompt } = require("../_shared/securePrompt");
const { MAX_ACCOUNTS, accountIndexForInput } = require("./store");

// The switcher at rest: a square in the window's bottom-left corner, the
// width of Outlook's left app bar. Outlook pages are never inset.
const SWITCHER_SIZE = 48;
// Window events after which the views must match the content size again.
// "resize" alone misses maximize/restore on some X11 setups
// (teams-for-linux #2947), and the first sizing when the window is shown
// (the content view's own bounds-changed covers that).
const RELAYOUT_EVENTS = ["resize", "maximize", "unmaximize", "restore", "enter-full-screen", "leave-full-screen", "show"];

/**
 * Multiple accounts in one window, ported from teams-for-linux
 * app/mainAppWindow/profileViewManager.js (its ADR-020): the main account is
 * the window's own page; every other account is a WebContentsView over it,
 * on its own session partition. All views stay attached and loaded; switching
 * shows one and hides the others, so nothing reloads and Outlook keeps
 * running in the background for every account (notifications, reminders,
 * unread counts). A view is never detached and attached again: on Linux a
 * view that loaded while detached paints but gets no input (teams-for-linux
 * #3057).
 *
 * The switcher (app/accountSwitcher) is one more transparent view, topmost,
 * in the bottom-left corner; it covers the window while its list is open.
 *
 * What makes a page an Outlook page (user agent, link policy, unread count,
 * remembered mailbox, ...) stays in app/mainAppWindow, which this calls
 * through `setupPage`.
 */
class AccountViews {
  #window;
  #accounts;
  #hooks;
  #webPreferences;
  #userAgent;
  #views = new Map(); // account id -> WebContentsView
  #switcher = null;
  #expanded = false;
  #unread = {};
  #relayout = () => this.#layout();

  /**
   * @param {Electron.BrowserWindow} window
   * @param {import("./store").AccountStore} accounts
   * @param {object} options
   * @param {(partition: string) => Electron.WebPreferences} options.webPreferences -
   *   the main window's, for another partition
   * @param {string} options.userAgent
   * @param {(contents: Electron.WebContents, account: object) => string} options.setupPage -
   *   makes a page an Outlook page; returns the URL to load
   * @param {(account: object) => void} [options.onRemoved]
   * @param {(account: object) => void} [options.onSwitched]
   */
  constructor(window, accounts, { webPreferences, userAgent, setupPage, onRemoved = () => {}, onSwitched = () => {} }) {
    this.#window = window;
    this.#accounts = accounts;
    this.#webPreferences = webPreferences;
    this.#userAgent = userAgent;
    this.#hooks = { setupPage, onRemoved, onSwitched };
  }

  initialize() {
    this.#registerIpc();
    for (const event of RELAYOUT_EVENTS) this.#window.on(event, this.#relayout);
    this.#window.contentView.on("bounds-changed", this.#relayout);
    // Keyboard focus goes back to the window's own page when the window is
    // focused again; give it to the visible account (teams-for-linux #3064).
    this.#window.on("focus", () => {
      if (!this.#expanded) this.activeContents?.focus();
    });
    // Electron mirrors only the window's own page's title onto the window.
    this.#window.on("page-title-updated", (event) => {
      if (!this.#accounts.getActive().main) event.preventDefault();
    });
    this.#window.once("closed", () => this.#dispose());
    this.#attachShortcuts(this.#window.webContents);

    for (const account of this.#accounts.list()) {
      if (!account.main) this.#createView(account);
    }
    this.#createSwitcher();
    this.#show(this.#accounts.getActive());
  }

  /** The visible account's Outlook page. */
  get activeContents() {
    return this.contentsOf(this.#accounts.getActive().id);
  }

  /**
   * @param {string} id
   * @returns {Electron.WebContents|null}
   */
  contentsOf(id) {
    const account = this.#accounts.get(id);
    if (!account) return null;
    const contents = account.main ? this.#window.webContents : this.#views.get(id)?.webContents;
    return contents && !contents.isDestroyed() ? contents : null;
  }

  /**
   * The account whose Outlook page this is (not its pop-outs).
   * @param {Electron.WebContents} contents
   * @returns {object|null}
   */
  accountOf(contents) {
    if (!contents) return null;
    if (contents === this.#window.webContents) return this.#accounts.list()[0];
    for (const [id, view] of this.#views) {
      if (view.webContents === contents) return this.#accounts.get(id);
    }
    return null;
  }

  /** @param {string} id */
  switch(id) {
    if (!this.#accounts.get(id)) return;
    this.#accounts.switch(id);
    this.#show(this.#accounts.get(id));
  }

  /** @param {number} index - position in the account list */
  switchToIndex(index) {
    const account = this.#accounts.list()[index];
    if (account) this.switch(account.id);
  }

  /** @param {Record<string, number>} counts - unread mail per account id */
  setUnread(counts) {
    this.#unread = counts ?? {};
    this.#pushState();
  }

  async add() {
    if (this.#accounts.list().length >= MAX_ACCOUNTS) {
      await dialog.showMessageBox(this.#window, {
        type: "info",
        message: `Outlook for Linux holds up to ${MAX_ACCOUNTS} accounts.`,
      });
      return;
    }
    const name = await this.#askName({
      title: "Add Account",
      heading: "Add an account",
      message: "A name for this account, such as Work or Personal. It then opens on Microsoft's sign-in page.",
      submitLabel: "Add",
      value: `Account ${this.#accounts.list().length + 1}`,
    });
    if (!name) return;
    const account = this.#accounts.add(name);
    console.info("[ACCOUNTS] Account added", { count: this.#accounts.list().length });
    this.#createView(account);
    this.switch(account.id);
  }

  async renameActive() {
    const account = this.#accounts.getActive();
    const name = await this.#askName({
      title: "Rename Account",
      heading: "Rename this account",
      message: "The name shown in the account switcher and on its notifications.",
      submitLabel: "Rename",
      value: account.name,
    });
    if (!name) return;
    this.#accounts.rename(account.id, name);
    this.#pushState();
    this.#hooks.onSwitched(this.#accounts.getActive());
  }

  async removeActive() {
    const account = this.#accounts.getActive();
    if (account.main) return;
    const { response } = await dialog.showMessageBox(this.#window, {
      type: "warning",
      buttons: ["Cancel", "Remove"],
      defaultId: 0,
      cancelId: 0,
      message: `Remove the account “${account.name}”?`,
      detail:
        "This signs it out of Outlook for Linux and deletes its data here (cookies, cached mail). Your mailbox itself is not affected; adding the account again needs a new sign-in.",
    });
    if (response !== 1) return;
    const { removed, activeId } = this.#accounts.remove(account.id);
    this.#destroyView(removed);
    this.#hooks.onRemoved(removed);
    console.info("[ACCOUNTS] Account removed", { count: this.#accounts.list().length });
    this.#show(this.#accounts.get(activeId));
  }

  // A name from a text prompt, or null when cancelled.
  async #askName(options) {
    try {
      const name = await showSecurePrompt({
        ...options,
        inputType: "text",
        placeholder: "Account name",
        parent: this.#window,
      });
      return name?.trim() || null;
    } catch {
      return null;
    }
  }

  #createView(account) {
    const view = new WebContentsView({
      webPreferences: {
        ...this.#webPreferences(account.partition),
        // Background accounts navigate too (sign-in, reloads); that must not
        // take keyboard focus from the visible one.
        focusOnNavigation: false,
      },
    });
    const contents = view.webContents;
    this.#views.set(account.id, view);
    const url = this.#hooks.setupPage(contents, account);
    this.#attachShortcuts(contents);
    contents.on("page-title-updated", (_event, title) => {
      if (this.#accounts.getActive().id === account.id) this.#window.setTitle(title);
    });
    contents.on("did-finish-load", () => {
      if (!this.#expanded && this.#accounts.getActive().id === account.id) contents.focus();
    });
    // The page can close itself; never leave a dead view on screen.
    contents.once("destroyed", () => {
      if (this.#views.get(account.id) !== view) return;
      this.#views.delete(account.id);
      this.#detach(view);
    });

    view.setBounds(this.#contentBounds());
    view.setVisible(false);
    this.#window.contentView.addChildView(view);
    this.#raiseSwitcher();
    contents.loadURL(url, { userAgent: this.#userAgent }).catch(() => {
      // ConnectionMonitor retries network failures; anything else shows
      // Chromium's error page.
    });
  }

  // Closes the view and deletes the account's data: re-adding it must not
  // find the previous sign-in.
  #destroyView(account) {
    const view = this.#views.get(account.id);
    this.#views.delete(account.id);
    if (view) {
      this.#detach(view);
      if (!view.webContents.isDestroyed()) view.webContents.close();
    }
    session
      .fromPartition(account.partition)
      .clearStorageData()
      .catch((err) => console.warn("[ACCOUNTS] Could not clear a removed account's data", { error: err.message }));
  }

  #detach(view) {
    try {
      this.#window.contentView.removeChildView(view);
    } catch {
      // the window is gone already
    }
  }

  #show(account) {
    for (const [id, view] of this.#views) view.setVisible(id === account.id);
    const view = this.#views.get(account.id);
    if (view) {
      view.setBounds(this.#contentBounds());
      // On an attached view, addChildView only moves it to the top.
      this.#window.contentView.addChildView(view);
    }
    this.#raiseSwitcher();
    const contents = this.contentsOf(account.id);
    if (contents) {
      contents.focus();
      this.#window.setTitle(contents.getTitle() || "Outlook for Linux");
    }
    this.#pushState();
    this.#hooks.onSwitched(account);
  }

  // Ctrl+Alt+1…9, read before the page sees the keys, as app/zoom does.
  #attachShortcuts(contents) {
    contents.on("before-input-event", (event, input) => {
      const index = accountIndexForInput(input);
      if (index === null) return;
      event.preventDefault();
      this.switchToIndex(index);
    });
  }

  #contentBounds() {
    const [width, height] = this.#window.getContentSize();
    return { x: 0, y: 0, width, height };
  }

  #layout() {
    if (this.#window.isDestroyed()) return;
    const bounds = this.#contentBounds();
    for (const view of this.#views.values()) view.setBounds(bounds);
    this.#layoutSwitcher();
  }

  // --- Switcher ---

  #createSwitcher() {
    const view = new WebContentsView({
      webPreferences: {
        preload: path.join(__dirname, "..", "accountSwitcher", "preload.js"),
        contextIsolation: true,
        sandbox: true,
        nodeIntegration: false,
      },
    });
    this.#switcher = view;
    view.setBackgroundColor("#00000000");
    this.#layoutSwitcher();
    this.#window.contentView.addChildView(view);
    this.#attachShortcuts(view.webContents);
    view.webContents.on("did-finish-load", () => this.#pushState());
    view.webContents.loadFile(path.join(__dirname, "..", "accountSwitcher", "switcher.html")).catch((err) => {
      console.error("[ACCOUNTS] The account switcher failed to load", { code: err.code ?? "unknown" });
    });
  }

  #layoutSwitcher() {
    if (!this.#switcher) return;
    const bounds = this.#contentBounds();
    if (this.#expanded) {
      this.#switcher.setBounds(bounds);
      return;
    }
    const size = Math.min(SWITCHER_SIZE, bounds.width, bounds.height);
    this.#switcher.setBounds({ x: 0, y: Math.max(0, bounds.height - size), width: size, height: size });
  }

  #raiseSwitcher() {
    if (!this.#switcher) return;
    try {
      this.#window.contentView.addChildView(this.#switcher);
    } catch {
      // the window is gone already
    }
  }

  #isSwitcher(event) {
    return Boolean(this.#switcher) && event.sender === this.#switcher.webContents;
  }

  #state() {
    const accounts = this.#accounts.list().map(({ id, name, initials, color, main }) => ({ id, name, initials, color, main }));
    return { accounts, activeId: this.#accounts.getActive().id, unread: this.#unread };
  }

  #pushState() {
    const contents = this.#switcher?.webContents;
    if (contents && !contents.isDestroyed()) contents.send("account-switcher:state", this.#state());
  }

  #registerIpc() {
    // The switcher's accounts, the visible one and unread counts (names and
    // initials only; nothing about the mailboxes)
    ipcMain.handle("account-switcher:state", (event) => (this.#isSwitcher(event) ? this.#state() : null));
    // Switch to the account the user picked in the switcher
    ipcMain.on("account-switcher:switch", (event, id) => {
      if (this.#isSwitcher(event) && typeof id === "string") this.switch(id);
    });
    // "Add account…" in the switcher
    ipcMain.on("account-switcher:add", (event) => {
      if (this.#isSwitcher(event)) this.#run(() => this.add());
    });
    // "Rename this account…" in the switcher
    ipcMain.on("account-switcher:rename", (event) => {
      if (this.#isSwitcher(event)) this.#run(() => this.renameActive());
    });
    // "Remove this account…" in the switcher
    ipcMain.on("account-switcher:remove", (event) => {
      if (this.#isSwitcher(event)) this.#run(() => this.removeActive());
    });
    // The switcher covers the window while its list is open (so a click
    // anywhere closes it) and shrinks back to its corner after
    ipcMain.handle("account-switcher:expand", (event, expanded) => {
      if (!this.#isSwitcher(event)) return;
      this.#expanded = expanded === true;
      this.#layoutSwitcher();
      this.#raiseSwitcher();
      if (!this.#expanded && this.#window.isFocused()) this.activeContents?.focus();
    });
  }

  // An unhandled rejection ends the app (see app/index.js).
  #run(action) {
    action().catch((err) => console.error("[ACCOUNTS] Account action failed", { error: err.message }));
  }

  #dispose() {
    for (const channel of ["account-switcher:state", "account-switcher:expand"]) ipcMain.removeHandler(channel);
    for (const channel of ["account-switcher:switch", "account-switcher:add", "account-switcher:rename", "account-switcher:remove"]) {
      ipcMain.removeAllListeners(channel);
    }
    this.#views.clear();
    this.#switcher = null;
  }
}

module.exports = AccountViews;
module.exports.SWITCHER_SIZE = SWITCHER_SIZE;
