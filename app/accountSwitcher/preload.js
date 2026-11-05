const { contextBridge, ipcRenderer } = require("electron");

// Bridge for the account switcher (app/accountSwitcher). It renders its own
// HTML and never touches Outlook's DOM, so it runs with contextIsolation and
// the sandbox, and reaches the main process only through these channels,
// which accept nothing from any other page.
contextBridge.exposeInMainWorld("accountSwitcher", {
  // { accounts: [{id, name, initials, color, main}], activeId, unread: {id: count} }
  state: () => ipcRenderer.invoke("account-switcher:state"),
  switch: (id) => ipcRenderer.send("account-switcher:switch", id),
  add: () => ipcRenderer.send("account-switcher:add"),
  rename: () => ipcRenderer.send("account-switcher:rename"),
  remove: () => ipcRenderer.send("account-switcher:remove"),
  // Grows the view over the window while the list is open (invoke, so the
  // list is shown only once the view is big enough not to clip it).
  setExpanded: (expanded) => ipcRenderer.invoke("account-switcher:expand", Boolean(expanded)),
  onState: (callback) => ipcRenderer.on("account-switcher:state", (_event, state) => callback(state)),
});
