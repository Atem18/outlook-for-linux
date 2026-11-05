const { test } = require("node:test");
const assert = require("node:assert/strict");
const { AccountStore, MAIN_ID, MAX_ACCOUNTS, cleanName, initialsOf, colorOf } = require("../../app/accounts/store");

function memoryStore(initial = {}) {
  const data = structuredClone(initial);
  return {
    get: (key) => structuredClone(data[key]),
    set: (key, value) => {
      data[key] = structuredClone(value);
    },
    data,
  };
}

test("a fresh store has only the main account, on the configured partition", () => {
  const accounts = new AccountStore(memoryStore(), "persist:outlook");
  const list = accounts.list();
  assert.equal(list.length, 1);
  assert.equal(list[0].id, MAIN_ID);
  assert.equal(list[0].main, true);
  assert.equal(list[0].partition, "persist:outlook");
  assert.equal(accounts.getActive().id, MAIN_ID);
});

test("added accounts get their own partition and do not become active", () => {
  const settings = memoryStore();
  const accounts = new AccountStore(settings, "persist:outlook");
  let changes = 0;
  accounts.on("change", () => changes++);
  const work = accounts.add("  Work   mail ");
  assert.equal(work.name, "Work mail");
  assert.equal(work.partition, `persist:outlook-account-${work.id}`);
  assert.equal(work.main, false);
  assert.equal(work.initials, "WM");
  assert.equal(accounts.getActive().id, MAIN_ID);
  assert.equal(changes, 1);
  // Survives a restart (a new store over the same settings).
  assert.deepEqual(new AccountStore(settings, "persist:outlook").get(work.id), work);
});

test("switching records the active account and emits it", () => {
  const accounts = new AccountStore(memoryStore(), "persist:outlook");
  const work = accounts.add("Work");
  let switched = null;
  accounts.on("switch", (account) => (switched = account));
  accounts.switch(work.id);
  assert.equal(accounts.getActive().id, work.id);
  assert.equal(switched.id, work.id);
  assert.throws(() => accounts.switch("nope"), /No such account/);
});

test("removing the visible account falls back to the main one; the main one stays", () => {
  const accounts = new AccountStore(memoryStore(), "persist:outlook");
  const work = accounts.add("Work");
  accounts.switch(work.id);
  const { removed, activeId } = accounts.remove(work.id);
  assert.equal(removed.partition, work.partition);
  assert.equal(activeId, MAIN_ID);
  assert.equal(accounts.list().length, 1);
  assert.throws(() => accounts.remove(MAIN_ID), /cannot be removed/);
  assert.throws(() => accounts.remove(work.id), /No such account/);
});

test("names are required, renames apply, and the count is capped", () => {
  const accounts = new AccountStore(memoryStore(), "persist:outlook");
  assert.throws(() => accounts.add("   "), /needs a name/);
  accounts.rename(MAIN_ID, "Personal");
  assert.equal(accounts.get(MAIN_ID).name, "Personal");
  assert.throws(() => accounts.rename(MAIN_ID, ""), /needs a name/);
  for (let i = 1; i < MAX_ACCOUNTS; i++) accounts.add(`Account ${i + 1}`);
  assert.throws(() => accounts.add("One too many"), /At most/);
});

test("the remembered mailbox is per account and can be forgotten", () => {
  const accounts = new AccountStore(memoryStore(), "persist:outlook");
  const work = accounts.add("Work");
  accounts.setMailboxOrigin(work.id, "https://outlook.office.com");
  assert.equal(accounts.get(work.id).mailboxOrigin, "https://outlook.office.com");
  accounts.setMailboxOrigin(work.id, null);
  assert.equal(accounts.get(work.id).mailboxOrigin, undefined);
});

test("a hand-edited store cannot point an account at another partition", () => {
  const settings = memoryStore({
    accounts: {
      list: [
        { id: "a", name: "Hijack", partition: "persist:outlook" },
        { id: "b", name: "Fine", partition: "persist:outlook-account-b", extra: "dropped" },
        { id: "b", name: "Duplicate", partition: "persist:outlook-account-b" },
        { id: "main", name: "Mine", partition: "persist:elsewhere" },
        "garbage",
      ],
      activeId: "a",
    },
  });
  const accounts = new AccountStore(settings, "persist:outlook");
  assert.deepEqual(
    accounts.list().map((account) => [account.id, account.name, account.partition]),
    [
      ["main", "Mine", "persist:outlook"],
      ["b", "Fine", "persist:outlook-account-b"],
    ]
  );
  assert.equal(accounts.getActive().id, MAIN_ID);
});

test("names, initials and colours", () => {
  assert.equal(cleanName(42), null);
  assert.equal(cleanName("x".repeat(100)).length, 64);
  assert.equal(initialsOf("contoso"), "CO");
  assert.equal(initialsOf("Jane van Doe"), "JD");
  assert.equal(initialsOf("😀 smile"), "😀S");
  assert.equal(initialsOf(""), "?");
  assert.equal(colorOf("x"), colorOf("x"));
  assert.match(colorOf("x"), /^hsl\(\d+, 55%, 42%\)$/);
});

test("the tray and dock show the sum of the accounts' unread counts", () => {
  const { unreadTotal } = require("../../app/tray/badge");
  assert.equal(unreadTotal([3, 0, 4]), 7);
  assert.equal(unreadTotal([]), 0);
  assert.equal(unreadTotal([2, -1, 1.5, Number.NaN, "3", null]), 2);
});

test("Ctrl+Alt+1…9 picks an account by position", () => {
  const { accountIndexForInput } = require("../../app/accounts/store");
  const key = (props) => accountIndexForInput({ type: "keyDown", control: true, alt: true, ...props });
  assert.equal(key({ code: "Digit1", key: "1" }), 0);
  assert.equal(key({ code: "Digit9", key: "9" }), 8);
  // AZERTY: the digit row types "&" without Shift.
  assert.equal(key({ code: "Digit1", key: "&" }), 0);
  assert.equal(key({ code: "Digit0", key: "0" }), null);
  assert.equal(key({ code: "KeyA", key: "a" }), null);
  assert.equal(key({ code: "Digit1", meta: true }), null);
  assert.equal(accountIndexForInput({ type: "keyDown", control: true, alt: false, code: "Digit1" }), null);
  assert.equal(accountIndexForInput({ type: "keyUp", control: true, alt: true, code: "Digit1" }), null);
});
