"use strict";

// When the main window's page fails to load because the network is down
// (startup without network, Wi-Fi dropping, resume from suspend), Chromium
// shows its error page and stays there. These decide when to try again.

// Chromium net error codes that mean "no network right now", as opposed to a
// real error worth showing (certificate, blocked request, HTTP failure).
const NETWORK_ERRORS = new Set([
  -7, // TIMED_OUT
  -21, // NETWORK_CHANGED
  -100, // CONNECTION_CLOSED
  -101, // CONNECTION_RESET
  -102, // CONNECTION_REFUSED
  -104, // CONNECTION_FAILED
  -105, // NAME_NOT_RESOLVED
  -106, // INTERNET_DISCONNECTED
  -109, // ADDRESS_UNREACHABLE
  -118, // CONNECTION_TIMED_OUT
  -137, // NAME_RESOLUTION_FAILED
]);

const BASE_DELAY_MS = 2_000;
const MAX_DELAY_MS = 60_000;

function isNetworkError(errorCode) {
  return NETWORK_ERRORS.has(errorCode);
}

// 2s, 4s, 8s ... capped at a minute.
function retryDelay(attempt) {
  const exponent = Math.max(0, Math.min(attempt, 10));
  return Math.min(BASE_DELAY_MS * 2 ** exponent, MAX_DELAY_MS);
}

module.exports = { isNetworkError, retryDelay, MAX_DELAY_MS };
