/**
 * Classifies error messages as transient network failures.
 *
 * The global uncaughtException / unhandledRejection handlers in app/index.js
 * use this to decide between "warn and keep running" and "log fatally and
 * exit". A dropped VPN or a suspended laptop must not kill the app.
 */

const { NETWORK_ERROR_PATTERNS } = require("../config/defaults");

function isNetworkError(message) {
  if (typeof message !== "string") return false;
  if (NETWORK_ERROR_PATTERNS.some((pattern) => message.includes(pattern))) {
    return true;
  }
  // "Object has been destroyed" occurs when the window is torn down during a
  // network-triggered operation (e.g. a reload after network recovery).
  if (message.includes("Object has been destroyed")) return true;
  // "Script failed to execute" occurs when executeJavaScript runs on a page
  // where the APIs are unavailable, e.g. a Chrome error page after
  // ERR_NAME_NOT_RESOLVED. A symptom of network failure, not a fatal error.
  if (message.includes("Script failed to execute")) return true;
  return false;
}

module.exports = { isNetworkError };
