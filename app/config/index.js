const yargs = require("yargs");
const fs = require("node:fs");
const path = require("node:path");
const configOptions = require("./options");

const SYSTEM_CONFIG_PATH = "/etc/outlook-for-linux/config.json";

function getConfigFilePath(configPath) {
  return path.join(configPath, "config.json");
}

// Returns the parsed file, {} when it does not exist, or throws on bad JSON.
function readJsonFile(filePath) {
  if (!fs.existsSync(filePath)) return {};
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

// System-wide config is the admin's baseline; the user file overrides it key
// by key. Shallow on purpose until nested options exist.
function loadConfigFiles(configPath, systemConfigPath = SYSTEM_CONFIG_PATH) {
  const result = { configFile: {}, error: null };

  let systemConfig = {};
  try {
    systemConfig = readJsonFile(systemConfigPath);
  } catch (e) {
    console.warn("[CONFIG] Ignoring invalid system-wide config file:", e.message);
  }

  let userConfig = {};
  try {
    userConfig = readJsonFile(getConfigFilePath(configPath));
  } catch (e) {
    result.error = e.message;
    console.warn("[CONFIG] Ignoring invalid user config file:", e.message);
  }

  result.configFile = { ...systemConfig, ...userConfig };
  return result;
}

function parseConfig(configFile, args, appVersion) {
  // yargs v18 requires explicit instantiation - it's no longer a singleton
  return yargs()
    .env("OUTLOOK_FOR_LINUX")
    .config(configFile)
    .version(appVersion)
    .options(configOptions)
    .strict(false)
    .help()
    .parse(args);
}

function argv(configPath, appVersion, args = process.argv.slice(1)) {
  const { configFile, error } = loadConfigFiles(configPath);
  const config = parseConfig(configFile, args, appVersion);
  if (error) config.error = error;
  console.info("[CONFIG] Loaded", { hasConfigFile: Object.keys(configFile).length > 0 });
  return config;
}

/**
 * Sets one option in the user config file, keeping the others, as
 * teams-for-linux's menus write theirs to its config store. Runs only on a
 * user action (App Icon menu): the startup config stays as loaded. A file
 * that is not valid JSON is never rewritten, so a typo cannot cost the
 * options in it; the error goes to the caller. Written to a temp file and
 * renamed, so a crash cannot leave half a file.
 * @param {string} configPath - Directory holding config.json (userData)
 * @param {string} name - option name
 * @param {unknown} value - undefined removes the option
 */
function writeConfigOption(configPath, name, value) {
  const file = getConfigFilePath(configPath);
  const config = readJsonFile(file);
  if (value === undefined) {
    if (!(name in config)) return;
    delete config[name];
  } else {
    config[name] = value;
  }
  const temp = `${file}.tmp-${process.pid}`;
  fs.mkdirSync(configPath, { recursive: true });
  fs.writeFileSync(temp, `${JSON.stringify(config, null, 2)}\n`);
  fs.renameSync(temp, file);
}

module.exports = argv;
module.exports.getConfigFilePath = getConfigFilePath;
module.exports.writeConfigOption = writeConfigOption;
module.exports.loadConfigFiles = loadConfigFiles;
