/**
 * Default values shared across modules. Kept free of Electron imports so unit
 * tests (and any future docs generator) can require it outside Electron.
 */

// Network error patterns that indicate transient connection issues
// (proxy, tunnel, DNS, etc.). Used by app/utils/networkErrors.js so the global
// error handlers do not terminate the app on a network blip.
const NETWORK_ERROR_PATTERNS = [
  "ERR_TUNNEL_CONNECTION_FAILED",
  "ERR_PROXY_CONNECTION_FAILED",
  "ERR_INTERNET_DISCONNECTED",
  "ERR_NETWORK_CHANGED",
  "ERR_CONNECTION_RESET",
  "ERR_CONNECTION_REFUSED",
  "ERR_CONNECTION_TIMED_OUT",
  "ERR_NAME_NOT_RESOLVED",
];

module.exports = { NETWORK_ERROR_PATTERNS };
