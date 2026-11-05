// The account switcher, ported from teams-for-linux app/profileSwitcher.
const pill = document.getElementById("pill");
const pillAvatar = document.getElementById("pill-avatar");
const pillUnread = document.getElementById("pill-unread");
const scrim = document.getElementById("scrim");
const dropdown = document.getElementById("dropdown");
const listEl = document.getElementById("account-list");
const removeButton = document.getElementById("remove");

const api = globalThis.accountSwitcher;
const DEFAULT_COLOR = "#0f6cbd";

let accounts = [];
let activeId = null;
let unread = {};
let open = false;

const active = () => accounts.find((account) => account.id === activeId) ?? null;

function unreadFor(id) {
  const count = unread[id];
  return Number.isInteger(count) && count > 0 ? count : 0;
}

function paintAvatar(el, account) {
  el.textContent = account?.initials || "?";
  // Only a valid CSS colour reaches the style.
  const color = account?.color;
  el.style.background = color && CSS.supports("background-color", color) ? color : DEFAULT_COLOR;
}

function renderPill() {
  const current = active();
  paintAvatar(pillAvatar, current);
  // The visible account's unread mail is on screen already; the dot is for
  // the others.
  const elsewhere = accounts.some((account) => account.id !== activeId && unreadFor(account.id) > 0);
  pillUnread.hidden = !elsewhere;
  const label = current ? `${current.name}: switch account` : "Switch account";
  pill.title = elsewhere ? `${label} (unread mail in another account)` : label;
  removeButton.disabled = !current || current.main;
  removeButton.title = current?.main ? "The first account is the one Outlook for Linux started with and stays" : "";
}

function renderList() {
  const focusedId = listEl.contains(document.activeElement) ? document.activeElement.dataset.accountId : null;
  listEl.textContent = "";
  accounts.forEach((account, index) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "account-row" + (account.id === activeId ? " active" : "");
    row.setAttribute("role", "menuitemradio");
    row.setAttribute("aria-checked", String(account.id === activeId));
    row.dataset.accountId = account.id;

    const avatar = document.createElement("span");
    avatar.className = "avatar";
    paintAvatar(avatar, account);
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = account.name;
    row.append(avatar, name);

    const count = unreadFor(account.id);
    if (count > 0) {
      const badge = document.createElement("span");
      badge.className = "unread-badge";
      badge.textContent = count > 99 ? "99+" : String(count);
      badge.setAttribute("aria-label", `${count} unread`);
      row.append(badge);
    }
    const hint = document.createElement("span");
    hint.className = "shortcut-hint";
    hint.textContent = `Ctrl+Alt+${index + 1}`;
    row.append(hint);

    row.addEventListener("click", () => choose(account.id));
    const li = document.createElement("li");
    li.append(row);
    listEl.append(li);
  });
  if (focusedId) listEl.querySelector(`[data-account-id="${CSS.escape(focusedId)}"]`)?.focus();
}

async function setOpen(next) {
  if (open === next) return;
  open = next;
  pill.setAttribute("aria-expanded", String(open));
  if (!open) {
    scrim.hidden = true;
    dropdown.hidden = true;
    api.setExpanded(false);
    return;
  }
  renderList();
  await api.setExpanded(true);
  if (!open) return;
  scrim.hidden = false;
  dropdown.hidden = false;
  listEl.querySelector(".account-row.active")?.focus();
}

function choose(id) {
  setOpen(false);
  if (id === activeId) return;
  activeId = id;
  renderPill();
  api.switch(id);
}

function applyState(state) {
  accounts = Array.isArray(state?.accounts) ? state.accounts : [];
  activeId = state?.activeId ?? null;
  unread = state?.unread && typeof state.unread === "object" ? state.unread : {};
  renderPill();
  if (open) renderList();
}

pill.addEventListener("click", () => setOpen(!open));
scrim.addEventListener("click", () => setOpen(false));
for (const [id, action] of [
  ["add", () => api.add()],
  ["rename", () => api.rename()],
  ["remove", () => api.remove()],
]) {
  document.getElementById(id).addEventListener("click", () => {
    setOpen(false);
    action();
  });
}
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") setOpen(false);
});
// Clicking into Outlook closes the list.
globalThis.addEventListener("blur", () => setOpen(false));

api.onState(applyState);
api.state().then(applyState, () => {});
