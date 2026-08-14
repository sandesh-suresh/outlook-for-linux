const yargs = require("yargs");
const fs = require("node:fs");
const path = require("node:path");
const logger = require("./logger");
const configOptions = require("./options");

function getConfigFilePath(configPath) {
  return path.join(configPath, "config.json");
}

/**
 * Reads the user's config file, if any.
 *
 * @returns {{ configFile: object, isConfigFile: boolean, configError: string|null }}
 */
function readConfigFile(configPath) {
  const filePath = getConfigFilePath(configPath);

  if (!fs.existsSync(filePath)) {
    console.info("[CONFIG] No config file found, using defaults");
    return { configFile: {}, isConfigFile: false, configError: null };
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8"));
    console.info("[CONFIG] Loaded user configuration");
    return { configFile: parsed, isConfigFile: true, configError: null };
  } catch (e) {
    // The message can contain a file path, so it is sanitized by the logger
    // hook before it reaches any transport.
    console.warn(`[CONFIG] Error in config file, using defaults: ${e.message}`);
    return { configFile: {}, isConfigFile: false, configError: e.message };
  }
}

/** Recursively freezes an object so the config is immutable after startup. */
function deepFreeze(obj) {
  for (const value of Object.values(obj)) {
    if (value && typeof value === "object" && !Object.isFrozen(value)) {
      deepFreeze(value);
    }
  }
  return Object.freeze(obj);
}

/**
 * Resolves the application configuration.
 *
 * Precedence: option defaults < config file < CLI flags.
 *
 * @param {string} configPath - Directory holding config.json.
 * @param {string} appVersion - Version reported by --version.
 * @param {string[]} [argvArray] - Argument vector; defaults to the real one.
 * @returns {object} The deep-frozen configuration.
 */
function argv(configPath, appVersion, argvArray = process.argv.slice(1)) {
  const { configFile, isConfigFile, configError } = readConfigFile(configPath);

  // yargs v18 is no longer a singleton and requires explicit instantiation.
  const config = yargs()
    .config(configFile)
    .version(appVersion)
    .options(configOptions)
    .help()
    .parse(argvArray);

  if (configError) {
    config.error = configError;
  }
  config.configPath = configPath;
  config.isConfigFile = isConfigFile;

  logger.init(config.logConfig);

  console.info("[CONFIG] Configuration resolved");

  return deepFreeze(config);
}

module.exports = argv;
