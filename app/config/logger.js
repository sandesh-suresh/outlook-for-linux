const { sanitizeLogData } = require("../utils/logSanitizer");

const FORBIDDEN_KEYS = new Set(["__proto__", "constructor", "prototype"]);

/**
 * Deep-merges source into target, preserving functions on the target by
 * assigning source properties onto them (electron-log's transports are
 * callable objects carrying configuration properties).
 */
function mergeWith(target, source) {
  for (const key of Object.keys(source)) {
    if (FORBIDDEN_KEYS.has(key)) continue;

    const targetValue = target[key];
    const sourceValue = source[key];

    if (typeof targetValue === "function") {
      Object.assign(targetValue, sourceValue);
    } else if (
      sourceValue &&
      typeof sourceValue === "object" &&
      !Array.isArray(sourceValue) &&
      targetValue &&
      typeof targetValue === "object" &&
      !Array.isArray(targetValue)
    ) {
      mergeWith(targetValue, sourceValue);
    } else {
      target[key] = sourceValue;
    }
  }
  return target;
}

function silenceConsole() {
  console.log = function () {};
  console.info = function () {};
  console.debug = function () {};
  console.warn = function () {};
  console.error = function () {};
}

/**
 * Initialises logging.
 *
 * @param {object|string|false} config - electron-log config, the literal
 *   string "console" to keep the plain console, or a falsy value to silence
 *   logging entirely.
 * @param {object} [log] - The electron-log instance. Injectable for tests.
 */
exports.init = function (config, log = require("electron-log/main")) {
  if (!config) {
    console.info("[LOGGER] Disabling logs");
    silenceConsole();
    return;
  }

  if (config === "console") {
    console.debug("[LOGGER] Using the default console");
    return;
  }

  mergeWith(log, config);
  log.initialize();

  // PII sanitisation applies to every transport, so nothing can leak through
  // a transport added later.
  log.hooks.push((message) => {
    message.data = sanitizeLogData(message.data);
    return message;
  });

  Object.assign(console, log.functions);
  console.debug("[LOGGER] Logger initialised");
};
