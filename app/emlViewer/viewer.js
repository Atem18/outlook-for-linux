// app/emlViewer/viewer.js: renders the message the main process parsed.
(async () => {
  const $ = (id) => document.getElementById(id);
  const message = await window.emlViewer.message();
  if (!message) return;

  $("subject").textContent = message.subject || "(no subject)";
  $("from").textContent = message.from.join(", ");

  let dateText = "";
  if (message.date) {
    const date = new Date(message.date);
    dateText = date.toLocaleString(undefined, { dateStyle: "full", timeStyle: "short" });
    $("date").textContent = dateText;
    $("date").dateTime = message.date;
  }

  for (const field of ["to", "cc"]) {
    if (!message[field].length) continue;
    $(field).textContent = message[field].join(", ");
    $(`${field}-row`).hidden = false;
  }

  const formatSize = (bytes) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };
  if (message.attachments.length) {
    const list = $("attachments");
    for (const attachment of message.attachments) {
      const item = document.createElement("li");
      const button = document.createElement("button");
      button.type = "button";
      button.className = "attachment";
      button.title = "Save attachment";
      const name = document.createElement("span");
      name.className = "attachment-name";
      name.textContent = attachment.filename;
      const size = document.createElement("span");
      size.className = "attachment-size";
      size.textContent = formatSize(attachment.size);
      button.append(name, size);
      button.addEventListener("click", () => window.emlViewer.saveAttachment(attachment.index));
      item.append(button);
      list.append(item);
    }
    list.hidden = false;
  }

  $("body").srcdoc = message.document;

  if (message.hasRemoteContent) {
    $("remote-banner").hidden = false;
    $("load-remote").addEventListener("click", async () => {
      const documentWithRemote = await window.emlViewer.remoteContent();
      if (documentWithRemote) $("body").srcdoc = documentWithRemote;
      $("remote-banner").hidden = true;
    });
  }

  $("reply").addEventListener("click", () => window.emlViewer.reply(dateText));
})();
