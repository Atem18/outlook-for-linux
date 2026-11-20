// app/emlViewer/viewerPreload.js: the viewer window's bridge. The main
// process answers only for the message it opened in this window.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("emlViewer", {
  message: () => ipcRenderer.invoke("eml-viewer:message"),
  remoteContent: () => ipcRenderer.invoke("eml-viewer:remote-content"),
  saveAttachment: (index) => {
    if (Number.isInteger(index)) return ipcRenderer.invoke("eml-viewer:save-attachment", index);
    return Promise.resolve(false);
  },
  reply: (dateText) => ipcRenderer.invoke("eml-viewer:reply", String(dateText ?? "")),
});
