#!/usr/bin/env node
// The server of the smoke test - one origin, laid out like an SAP system:
//
//   /sap/bc/z2ui5                         abap2UI5 (@abap2ui5/node-runtime)
//                                         with the demo app of src/01 on top
//                                         (npm run e2e:build) - the page, the
//                                         roundtrips and ?z2ui5-bundle
//   /sap/bc/ui5_ui5/sap/z2ui5_xl/         the BSP of src/02 AS IT IS DELIVERED:
//                                         every page read from its abapGit
//                                         file, the 255-character padding cut
//   /sap/bc/ui5_ui5/sap/z2ui5_xl/resources/
//                                         UI5 from the @openui5 npm packages -
//                                         what the system's UI5 is under a BSP
//                                         (--ui5 1.136, the default, or 1.71)
//
// Office.js is not served: it comes from Microsoft's CDN, and the tests
// answer that URL with test/e2e/office-stub.js.
//
// ExcelBridge, until the runtime has it: z2ui5.cc.ExcelBridge is new in
// abap2UI5 (#2847) and not in the frontend of @abap2ui5/node-runtime
// 1.146.0. When the bundle does not define the module, UI5 asks for it as a
// file - resources/z2ui5/cc/ExcelBridge.js - and this server answers with
// the copy in test/e2e/vendor/ (its header names the commit). With a
// runtime whose bundle has the control the copy is never asked for, and
// the server says which of the two the run used.
//
//   node test/e2e/server.mjs [--port 4310] [--ui5 1.136|1.71] [--csp]
//
// --csp sends the Content-Security-Policy the README recommends for the
// host page with every page of the BSP (the tests use it).

import express from "express";
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { initialize, createApp } from "@abap2ui5/node-runtime";

const ROOT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
);
const require = createRequire(import.meta.url);
const arg = (name) => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const port = Number(arg("port") || process.env.PORT || 4310);
const host = arg("host") || "127.0.0.1";
const ui5 = arg("ui5") || "1.136";
const csp = process.argv.includes("--csp");

export const BSP_PATH = "/sap/bc/ui5_ui5/sap/z2ui5_xl";
export const ENDPOINT = "/sap/bc/z2ui5";

// The policy of README "Content-Security-Policy", for a test origin: the
// Office domains in frame-ancestors are the ones of Excel on the web.
// 'unsafe-eval' is there for UI5 from SOURCE files (the npm packages), which
// evaluates what a built UI5 has preloaded; the README says when a system
// needs it.
export const POLICY = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-eval' https://officeapis.public.onecdn.static.microsoft https://appsforoffice.microsoft.com",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data:",
  "font-src 'self' data:",
  "connect-src 'self'",
  "frame-ancestors 'self' https://*.cloud.microsoft https://*.officeapps.live.com https://*.office.com https://*.office365.com https://*.sharepoint.com",
].join("; ");

// --- UI5 from the npm packages -------------------------------------------
function ui5Roots(version) {
  const anchor =
    version === "1.71"
      ? path.dirname(require.resolve("ui5-1.71-sap.m/package.json"))
      : path.dirname(require.resolve("@openui5/sap.m/package.json"));
  const roots = [path.join(anchor, "src")];
  // the libraries sap.m depends on: next to it, or nested below it
  for (const base of [
    path.join(anchor, "node_modules", "@openui5"),
    path.dirname(anchor),
  ]) {
    if (!fs.existsSync(base)) continue;
    for (const name of fs.readdirSync(base)) {
      const src = path.join(base, name, "src");
      if (
        name.startsWith("sap.") &&
        fs.existsSync(src) &&
        !roots.includes(src)
      ) {
        const pkg = JSON.parse(
          fs.readFileSync(path.join(base, name, "package.json"), "utf8"),
        );
        const core = JSON.parse(
          fs.readFileSync(path.join(anchor, "package.json"), "utf8"),
        );
        if (pkg.version === core.version) roots.push(src);
      }
    }
  }
  return roots;
}

const ROOTS = ui5Roots(ui5);
const MIME = {
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".properties": "text/plain; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".png": "image/png",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ttf": "font/ttf",
  ".woff2": "font/woff2",
};

// sap-ui-version.json as a UI5 build writes it - the source packages carry
// none, and the frontend reads the UI5 version from it
function versionInfo() {
  const libraries = ROOTS.map((root) => {
    const pkg = JSON.parse(
      fs.readFileSync(path.join(root, "..", "package.json"), "utf8"),
    );
    return { name: pkg.name.replace(/^@openui5\//, ""), version: pkg.version };
  });
  const version = libraries.find((l) => l.name === "sap.m").version;
  return {
    name: "OpenUI5 (npm source packages)",
    version,
    buildTimestamp: "202601010000",
    gav: `com.sap.openui5.dist:sdk:${version}`,
    libraries: libraries.map((l) => ({ ...l, buildTimestamp: "202601010000" })),
  };
}

// sap-ui-core.js of the SOURCE packages loads the loader and the core with
// document.write( ) - and two of the scripts it writes are inline, which a
// policy without 'unsafe-inline' refuses. A built UI5 (the one of a system,
// the CDN) has no such thing: its sap-ui-core.js is one file of the loader,
// its configuration and the core. This is that file, made from the sources:
// the loader and its configuration as they are, then the two calls the
// inline scripts make.
function coreBootstrap() {
  const core = ROOTS.find((root) =>
    fs.existsSync(path.join(root, "ui5loader.js")),
  );
  return [
    fs.readFileSync(path.join(core, "ui5loader.js"), "utf8"),
    fs.readFileSync(path.join(core, "ui5loader-autoconfig.js"), "utf8"),
    'sap.ui.requireSync("sap/ui/core/Core");',
    "sap.ui.getCore().boot && sap.ui.getCore().boot();",
    "",
  ].join("\n;\n");
}

const VENDORED_BRIDGE = path.join(
  ROOT,
  "test",
  "e2e",
  "vendor",
  "ExcelBridge.js",
);
export const bridgeSource = { used: "bundle" };

function serveUi5(req, res) {
  const rel = decodeURIComponent(req.path.replace(/^\/+/, "")).replace(
    /^sap-ui-cachebuster\//,
    "",
  );
  if (rel === "sap-ui-version.json") return res.json(versionInfo());
  if (rel === "sap-ui-core.js")
    return res.type(MIME[".js"]).send(coreBootstrap());
  if (rel === "z2ui5/cc/ExcelBridge.js") {
    bridgeSource.used = "vendored";
    console.log(
      "server: z2ui5/cc/ExcelBridge.js is not in the bundle - the vendored copy answers",
    );
    return res.type(MIME[".js"]).send(fs.readFileSync(VENDORED_BRIDGE));
  }
  for (const root of ROOTS) {
    const full = path.join(root, rel);
    if (
      full.startsWith(root + path.sep) &&
      fs.existsSync(full) &&
      fs.statSync(full).isFile()
    ) {
      return res
        .type(MIME[path.extname(full)] || "application/octet-stream")
        .send(fs.readFileSync(full));
    }
  }
  return res.status(404).send("not found");
}

// --- the BSP as delivered -------------------------------------------------
function bspPages() {
  const dir = path.join(ROOT, "src", "02");
  const wapa = fs.readFileSync(path.join(dir, "z2ui5_xl.wapa.xml"), "utf8");
  const pages = new Map();
  for (const m of wapa.matchAll(/<PAGENAME>([^<]+)<\/PAGENAME>/g)) {
    const name = m[1].replace(/&amp;/g, "&");
    const file = path.join(
      dir,
      `z2ui5_xl.wapa.${name.replace(/\//g, "_-").toLowerCase()}`,
    );
    const text = fs
      .readFileSync(file, "utf8")
      .split("\n")
      .map((line) => line.replace(/ +$/, ""))
      .join("\n");
    pages.set(name, text + "\n");
  }
  return pages;
}

const PAGES = bspPages();

function serveBsp(req, res) {
  const name = decodeURIComponent(req.path.replace(/^\/+/, "")) || "index.html";
  const page = PAGES.get(name);
  if (page === undefined) return res.status(404).send("not found");
  if (csp) res.set("Content-Security-Policy", POLICY);
  res.set("Cache-Control", "no-store");
  return res.type(MIME[path.extname(name)] || "text/plain").send(page);
}

// --- the backend ------------------------------------------------------------
const apps = path.join(ROOT, "test", "e2e", ".build", "apps", "index.mjs");
if (!fs.existsSync(apps)) {
  console.error(
    "server: test/e2e/.build/apps is missing - run npm run e2e:build first",
  );
  process.exit(1);
}
await initialize();
await import(pathToFileURL(apps).href);

const app = express();
app.use(ENDPOINT, await createApp());
app.use(`${BSP_PATH}/resources`, serveUi5);
app.use(BSP_PATH, serveBsp);
app.get("/", (_req, res) => res.redirect(`${BSP_PATH}/index.html`));

const server = app.listen(port, host, () => {
  const runtime = require("@abap2ui5/node-runtime/package.json").version;
  console.log(
    `server: http://${host}:${port}${BSP_PATH}/index.html?app=Z2UI5_CL_XL_DEMO ` +
      `(UI5 ${versionInfo().version}, @abap2ui5/node-runtime ${runtime}${csp ? ", CSP" : ""})`,
  );
});
const stop = () => server.close(() => process.exit(0));
process.on("SIGTERM", stop);
process.on("SIGINT", stop);
