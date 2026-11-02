// Bundles app/browser/preload.js and its tools into one file.
//
// The main window's preload runs sandboxed: it can require "electron" but not
// local files, so the tools have to be inlined. Run by `npm start`, the test
// and packaging scripts; the output is git-ignored.
const path = require("node:path");
const esbuild = require("esbuild");

const root = path.join(__dirname, "..");

esbuild
  .build({
    entryPoints: [path.join(root, "app", "browser", "preload.js")],
    outfile: path.join(root, "app", "browser", "preload.bundle.js"),
    bundle: true,
    platform: "browser",
    format: "cjs",
    target: "chrome140",
    external: ["electron"],
    logLevel: "warning",
  })
  .catch(() => process.exit(1));
