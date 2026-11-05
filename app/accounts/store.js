const { EventEmitter } = require("node:events");
const crypto = require("node:crypto");

/**
 * The accounts of the multiple-accounts mode (`multiAccount`), kept in the
 * settings store. Ported from teams-for-linux app/profilesManager, reduced to
 * what Outlook needs (no pinning, URL override or per-profile mute), and
 * Electron-free so node:test covers it.
 *
 * The first account is the main window's own session (config.partition), so
 * turning the mode on keeps the current sign-in. It has the fixed id "main"
 * and cannot be removed. Every other account gets its own partition,
 * `persist:outlook-account-<uuid>`, fixed for its lifetime.
 *
 * Settings shape under `accounts`:
 *   { list: [{ id, name, partition, mailboxOrigin? }], activeId }
 * The main account's partition is stored as null: it follows config.partition.
 *
 * Events: "change" (after any add/rename/remove), "switch" (account).
 */

const STORE_KEY = "accounts";
const MAIN_ID = "main";
const PARTITION_PREFIX = "persist:outlook-account-";
const MAX_ACCOUNTS = 9; // one Ctrl+Alt+<digit> shortcut each
const MAX_NAME_LENGTH = 64;
const DEFAULT_MAIN_NAME = "Account 1";

/**
 * A display name as typed: trimmed, inner whitespace collapsed, capped.
 * @param {unknown} name
 * @returns {string|null} null when nothing is left
 */
function cleanName(name) {
  if (typeof name !== "string") return null;
  const cleaned = Array.from(name.trim().replaceAll(/\s+/g, " ")).slice(0, MAX_NAME_LENGTH).join("");
  return cleaned || null;
}

/**
 * Two letters for the account's avatar: the first letters of the first and
 * last words, or the first two of a single word. By code point, so an emoji
 * or a supplementary-plane letter is not split.
 * @param {string} name
 * @returns {string}
 */
function initialsOf(name) {
  const words = (name || "").split(/\s+/).filter(Boolean);
  if (words.length === 0) return "?";
  if (words.length === 1) return Array.from(words[0]).slice(0, 2).join("").toLocaleUpperCase();
  return (Array.from(words[0])[0] + Array.from(words.at(-1))[0]).toLocaleUpperCase();
}

/**
 * A stable avatar colour from the account id, dark enough for white text.
 * @param {string} id
 * @returns {string} CSS hsl() colour
 */
function colorOf(id) {
  const hue = Math.round((crypto.createHash("sha256").update(String(id)).digest()[0] * 360) / 256);
  return `hsl(${hue}, 55%, 42%)`;
}

class AccountStore {
  #settingsStore;
  #mainPartition;
  #emitter = new EventEmitter();

  /**
   * @param {{get: Function, set: Function}} settingsStore
   * @param {string} mainPartition - config.partition, the main window's
   */
  constructor(settingsStore, mainPartition) {
    this.#settingsStore = settingsStore;
    this.#mainPartition = mainPartition;
  }

  on(event, handler) {
    this.#emitter.on(event, handler);
  }

  /**
   * Every account, the main one first, each with its resolved partition,
   * initials and colour.
   * @returns {Array<{id: string, name: string, partition: string, main: boolean, initials: string, color: string, mailboxOrigin?: string}>}
   */
  list() {
    return this.#read().list.map((record) => this.#view(record));
  }

  /** @param {string} id */
  get(id) {
    const record = this.#read().list.find((account) => account.id === id);
    return record ? this.#view(record) : null;
  }

  /** The visible account; the main one when none is recorded. */
  getActive() {
    const { list, activeId } = this.#read();
    return this.#view(list.find((account) => account.id === activeId) ?? list[0]);
  }

  /**
   * @param {string} id
   * @returns {object} the account now active
   */
  switch(id) {
    const state = this.#read();
    const record = state.list.find((account) => account.id === id);
    if (!record) throw new Error("No such account");
    if (state.activeId !== id) {
      state.activeId = id;
      this.#write(state);
    }
    const account = this.#view(record);
    this.#emitter.emit("switch", account);
    return account;
  }

  /**
   * Adds an account with a fresh partition. Does not switch to it.
   * @param {string} name
   * @returns {object} the new account
   */
  add(name) {
    const cleaned = cleanName(name);
    if (!cleaned) throw new Error("An account needs a name");
    const state = this.#read();
    if (state.list.length >= MAX_ACCOUNTS) throw new Error(`At most ${MAX_ACCOUNTS} accounts`);
    const id = crypto.randomUUID();
    const record = { id, name: cleaned, partition: `${PARTITION_PREFIX}${id}` };
    state.list.push(record);
    this.#write(state);
    const account = this.#view(record);
    this.#emitter.emit("change");
    return account;
  }

  /**
   * @param {string} id
   * @param {string} name
   */
  rename(id, name) {
    const cleaned = cleanName(name);
    if (!cleaned) throw new Error("An account needs a name");
    this.#patch(id, { name: cleaned });
    this.#emitter.emit("change");
  }

  /**
   * Remembers which mailbox an account ends up in (personal or work), as
   * app/mainAppWindow/startUrl.js does for the main account.
   * @param {string} id
   * @param {string|null} origin
   */
  setMailboxOrigin(id, origin) {
    this.#patch(id, { mailboxOrigin: origin ?? undefined });
  }

  /**
   * Removes an account other than the main one. Its storage is the caller's
   * to clear. When it was visible, the main account becomes active.
   * @param {string} id
   * @returns {{removed: object, activeId: string}}
   */
  remove(id) {
    if (id === MAIN_ID) throw new Error("The main account cannot be removed");
    const state = this.#read();
    const index = state.list.findIndex((account) => account.id === id);
    if (index === -1) throw new Error("No such account");
    const [record] = state.list.splice(index, 1);
    if (state.activeId === id) state.activeId = MAIN_ID;
    this.#write(state);
    this.#emitter.emit("change");
    return { removed: this.#view(record), activeId: state.activeId };
  }

  #patch(id, fields) {
    const state = this.#read();
    const record = state.list.find((account) => account.id === id);
    if (!record) throw new Error("No such account");
    Object.assign(record, fields);
    for (const [key, value] of Object.entries(fields)) {
      if (value === undefined) delete record[key];
    }
    this.#write(state);
  }

  #view(record) {
    const main = record.id === MAIN_ID;
    const view = {
      id: record.id,
      name: record.name,
      partition: main ? this.#mainPartition : record.partition,
      main,
      initials: initialsOf(record.name),
      color: colorOf(record.id),
    };
    if (record.mailboxOrigin) view.mailboxOrigin = record.mailboxOrigin;
    return view;
  }

  // Tolerates a missing or hand-edited entry: unknown fields and malformed
  // records are dropped, and the main account is always first.
  #read() {
    const raw = this.#settingsStore.get(STORE_KEY);
    const stored = Array.isArray(raw?.list) ? raw.list : [];
    const seen = new Set([MAIN_ID]);
    const others = [];
    for (const record of stored) {
      const valid =
        record &&
        typeof record.id === "string" &&
        !seen.has(record.id) &&
        typeof record.partition === "string" &&
        record.partition === `${PARTITION_PREFIX}${record.id}` &&
        cleanName(record.name);
      if (!valid) continue;
      seen.add(record.id);
      const clean = { id: record.id, name: cleanName(record.name), partition: record.partition };
      if (typeof record.mailboxOrigin === "string") clean.mailboxOrigin = record.mailboxOrigin;
      others.push(clean);
    }
    const storedMain = stored.find((record) => record?.id === MAIN_ID);
    const main = { id: MAIN_ID, name: cleanName(storedMain?.name) ?? DEFAULT_MAIN_NAME, partition: null };
    const list = [main, ...others.slice(0, MAX_ACCOUNTS - 1)];
    const activeId = list.some((account) => account.id === raw?.activeId) ? raw.activeId : MAIN_ID;
    return { list, activeId };
  }

  #write(state) {
    this.#settingsStore.set(STORE_KEY, state);
  }
}

/**
 * The account a Ctrl+Alt+<digit> key press switches to, by position (the
 * digit key's code, so it works on layouts where digits need Shift).
 * @param {{type?: string, code?: string, control?: boolean, alt?: boolean, meta?: boolean, shift?: boolean}} input -
 *   Electron's before-input-event input
 * @returns {number|null} index into the account list, or null
 */
function accountIndexForInput(input) {
  if (input?.type !== "keyDown" || !input.control || !input.alt || input.meta) return null;
  const match = /^Digit([1-9])$/.exec(input.code ?? "");
  return match ? Number(match[1]) - 1 : null;
}

module.exports = {
  AccountStore,
  MAIN_ID,
  MAX_ACCOUNTS,
  PARTITION_PREFIX,
  cleanName,
  initialsOf,
  colorOf,
  accountIndexForInput,
};
