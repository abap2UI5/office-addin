#!/usr/bin/env node
// Builds the backend of the smoke test: the demo app (src/01) transpiled
// against @abap2ui5/node-runtime, the recipe of that package's README
// ("Your own apps"):
//
//   1. open-abap-core at the commit the runtime was built with (its
//      package.json, abap2ui5.openAbapCore) -> test/e2e/.deps/open-abap-core
//      (OPEN_ABAP_CORE_DIR=<a checkout at that commit> skips the fetch)
//   2. abap_transpile, in the version the runtime names, of src/01 ->
//      test/e2e/.build/output
//   3. abap2ui5-own-apps -> test/e2e/.build/apps: the class alone, its
//      imports pointed at the runtime's own modules
//
// test/e2e/server.mjs imports test/e2e/.build/apps/index.mjs after the boot.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
);
const require = createRequire(import.meta.url);
const runtimePkg = require("@abap2ui5/node-runtime/package.json");
const { transpiler, openAbapCore } = runtimePkg.abap2ui5;
const installed = require("@abaplint/transpiler-cli/package.json").version;
if (installed !== transpiler) {
  console.error(
    `build-backend: @abaplint/transpiler-cli ${installed} is installed, the ` +
      `runtime ${runtimePkg.version} was transpiled with ${transpiler} - pin that one`,
  );
  process.exit(1);
}

const deps = path.join(ROOT, "test", "e2e", ".deps", "open-abap-core");
const build = path.join(ROOT, "test", "e2e", ".build");
const run = (cmd, args) =>
  execFileSync(cmd, args, { cwd: ROOT, stdio: "inherit" });

function headOf(dir) {
  try {
    return execFileSync("git", ["-C", dir, "rev-parse", "HEAD"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "ignore"],
    }).trim();
  } catch {
    return null;
  }
}

if (headOf(deps) !== openAbapCore) {
  fs.rmSync(deps, { recursive: true, force: true });
  fs.mkdirSync(path.dirname(deps), { recursive: true });
  const given = process.env.OPEN_ABAP_CORE_DIR;
  if (given) {
    if (!fs.existsSync(path.join(given, "src"))) {
      console.error(
        `build-backend: OPEN_ABAP_CORE_DIR=${given} is no open-abap-core checkout`,
      );
      process.exit(1);
    }
    fs.cpSync(given, deps, { recursive: true });
  } else {
    run("git", ["init", "-q", deps]);
    run("git", [
      "-C",
      deps,
      "fetch",
      "-q",
      "--depth",
      "1",
      "https://github.com/open-abap/open-abap-core",
      openAbapCore,
    ]);
    run("git", ["-C", deps, "checkout", "-q", "FETCH_HEAD"]);
  }
}

fs.rmSync(build, { recursive: true, force: true });
fs.mkdirSync(build, { recursive: true });
const config = {
  input_folder: "src/01",
  output_folder: "test/e2e/.build/output",
  libs: [
    {
      folder: "/node_modules/@abap2ui5/node-runtime/downport",
      files: "/**/*.*",
    },
    {
      url: "https://github.com/open-abap/open-abap-core",
      folder: "/test/e2e/.deps/open-abap-core",
    },
  ],
  write_unit_tests: false,
  options: {
    ignoreSyntaxCheck: false,
    addFilenames: true,
    addCommonJS: true,
    unknownTypes: "runtimeError",
  },
};
const configFile = path.join(build, "abap_transpile.json");
fs.writeFileSync(configFile, JSON.stringify(config, null, 2));
const cli = path.dirname(
  require.resolve("@abaplint/transpiler-cli/package.json"),
);
run(process.execPath, [
  path.join(cli, "abap_transpile"),
  path.relative(ROOT, configFile),
]);
const { ownApps } = await import("@abap2ui5/node-runtime/setup/own-apps.mjs");
const res = ownApps({
  output: path.join(build, "output"),
  apps: path.join(build, "apps"),
});
fs.rmSync(path.join(build, "output"), { recursive: true, force: true });
console.log(
  `build-backend: ${res.modules.length} app module(s) in test/e2e/.build/apps ` +
    `for @abap2ui5/node-runtime ${runtimePkg.version}`,
);
