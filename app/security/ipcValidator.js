/**
 * IPC security validation.
 *
 * The preload script shares the Outlook page's context, so the main process
 * treats every inbound channel name as untrusted: it must appear on this
 * allowlist, and its payload is stripped of prototype-pollution vectors.
 *
 * Add a channel here in the same change that registers its handler.
 */
const allowedChannels = new Set([
  // Renderer asks main for the resolved startup config (invoke).
  "get-config",

  // Notifications: renderer forwards Outlook's web notification (invoke),
  // main tells the renderer the OS dismissed it (main → renderer; not gated
  // by this validator, listed so the allowlist stays authoritative).
  "show-notification",
  "notification-closed",

  // Tray: renderer pushes a rendered badge icon (send) and the unread count
  // for the OS badge (invoke).
  "tray-update",
  "set-badge-count",

  // Renderer-side error forwarding, registered in app/browser/preload.js.
  "unhandled-rejection",
  "window-error",
]);

const DANGEROUS_PROPS = new Set(["__proto__", "constructor", "prototype"]);
const MAX_SANITIZE_DEPTH = 10;

/**
 * Recursively removes prototype-pollution vectors from a payload, in place.
 *
 * @param {any} obj - The payload to sanitize.
 * @param {number} depth - Recursion depth; bounded so cycles cannot overflow.
 */
function sanitizePayload(obj, depth = 0) {
  if (!obj || typeof obj !== "object" || depth > MAX_SANITIZE_DEPTH) {
    return;
  }

  for (const prop of DANGEROUS_PROPS) {
    if (Object.hasOwn(obj, prop)) {
      delete obj[prop];
    }
  }

  for (const key of Object.keys(obj)) {
    if (obj[key] && typeof obj[key] === "object") {
      sanitizePayload(obj[key], depth + 1);
    }
  }
}

/**
 * Validates an inbound IPC request.
 *
 * @param {string} channel - The IPC channel name.
 * @param {any} [payload] - The payload, sanitized in place when present.
 * @returns {boolean} True when the request may proceed.
 */
function validateIpcChannel(channel, payload = null) {
  if (!allowedChannels.has(channel)) {
    console.warn(`[IPC] Blocked unauthorized channel: ${channel}`);
    return false;
  }

  sanitizePayload(payload);
  return true;
}

module.exports = { validateIpcChannel, allowedChannels };
