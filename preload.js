'use strict';

/**
 * preload.js
 *
 * Lightweight bridge — just watches document.title and forwards
 * tray-update IPC when it changes.
 *
 * The authoritative unread count comes from main.js via
 * webContents.executeJavaScript() which runs in the page's main world
 * and can query the Outlook Web DOM directly.
 * This preload fires as a fast-path complement when the title DOES
 * carry the count (it sometimes does, e.g. when the window is hidden).
 */

const { ipcRenderer } = require('electron');

const TITLE_REGEX = /^\((\d+)\)/;

function extractFromTitle(title) {
  const m = (title || '').match(TITLE_REGEX);
  return m ? parseInt(m[1], 10) : -1; // -1 = "not in title"
}

let lastSent = -1;

function checkTitle() {
  const count = extractFromTitle(document.title);
  if (count >= 0 && count !== lastSent) {
    lastSent = count;
    ipcRenderer.send('tray-update', { count, source: 'title' });
  }
}

function watchTitle() {
  checkTitle();
  const titleEl = document.querySelector('title');
  if (titleEl) {
    new MutationObserver(checkTitle).observe(titleEl, {
      childList: true, characterData: true, subtree: true,
    });
  }
  new MutationObserver(checkTitle).observe(document.head, { childList: true });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', watchTitle);
} else {
  watchTitle();
}
