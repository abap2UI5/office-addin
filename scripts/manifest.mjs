#!/usr/bin/env node
// Writes the Office add-in manifest (XML, "add-in only manifest") for one
// SAP system from manifest/manifest.template.xml:
//
//   npm run manifest -- --host https://my.sap.example
//   npm run manifest -- --host https://my.sap.example \
//     --path /sap/bc/ui5_ui5/sap/z2ui5_xl/index.html --app Z2UI5_CL_MY_APP \
//     --client 100 --out manifest.xml
//
// Options (README, "The manifest"):
//   --host <origin>        the SAP system as the browser reaches it - https,
//                          no path (required)
//   --path <path>          the host page on it
//                          (default /sap/bc/ui5_ui5/sap/z2ui5_xl/index.html)
//   --app <class>          the abap2UI5 app the task pane starts
//                          (default Z2UI5_CL_XL_DEMO)
//   --endpoint <path>      the abap2UI5 service, when not /sap/bc/z2ui5
//   --client <nnn>         sap-client for the host page
//   --language <xx>        sap-language for the host page
//   --theme <theme>        sap-ui-theme for the host page (sap_fiori_3 on a
//                          system below UI5 1.102, which has no sap_horizon)
//   --param <name=value>   a startup parameter for the app (repeatable)
//   --app-domain <origin>  one more domain the task pane may navigate to -
//                          the identity provider of a SAML logon (repeatable)
//   --icons <url>          the folder the icons are served from - every
//                          icon-<size>.png of manifest/icons (default: this
//                          repository on GitHub)
//   --id <guid>            the add-in id (default: derived from host, path
//                          and app, so a manifest written again keeps it)
//   --version <n.n.n.n>    (default 1.0.0.0) - raise it with every change of
//                          a deployed manifest
//   --name, --description, --provider, --button, --support-url
//   --out <file>           (default manifest.xml; - writes to stdout)

import { createHash } from "node:crypto";
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
export const TEMPLATE = join(root, "manifest", "manifest.template.xml");

export const DEFAULTS = {
  path: "/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html",
  app: "Z2UI5_CL_XL_DEMO",
  version: "1.0.0.0",
  name: "abap2UI5 for Excel",
  description: "abap2UI5 apps of your SAP system next to the workbook.",
  provider: "abap2UI5",
  button: "abap2UI5",
  supportUrl: "https://github.com/abap2UI5/office-addin",
  icons:
    "https://raw.githubusercontent.com/abap2UI5/office-addin/main/manifest/icons/",
  out: "manifest.xml",
};

export const ICON_SIZES = [16, 32, 64, 80];
const CLASS_NAME = /^(\/[A-Z0-9_]{1,10}\/)?[A-Z0-9_]{1,30}$/;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const VERSION = /^\d{1,5}(\.\d{1,5}){0,3}$/;
// the length limits of the manifest schema: ShortString 125, LongString 250
const SHORT = 125;
const LONG = 250;

const MULTI = new Set(["param", "app-domain"]);
const KNOWN = new Set([
  "host",
  "path",
  "app",
  "endpoint",
  "client",
  "language",
  "theme",
  "param",
  "app-domain",
  "icons",
  "id",
  "version",
  "name",
  "description",
  "provider",
  "button",
  "support-url",
  "out",
]);

export function parseArgs(argv) {
  const opts = { param: [], "app-domain": [] };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (!a.startsWith("--")) throw new Error(`unexpected argument '${a}'`);
    let name = a.slice(2);
    let value;
    const eq = name.indexOf("=");
    if (eq >= 0) {
      value = name.slice(eq + 1);
      name = name.slice(0, eq);
    } else {
      value = argv[i + 1];
      i += 1;
    }
    if (!KNOWN.has(name)) throw new Error(`unknown option --${name}`);
    if (value === undefined) throw new Error(`--${name} needs a value`);
    if (MULTI.has(name)) opts[name].push(value);
    else opts[name] = value;
  }
  return opts;
}

function httpsOrigin(value, what) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${what} '${value}' is no URL`);
  }
  if (url.protocol !== "https:") {
    throw new Error(
      `${what} '${value}' is not https - Office loads add-ins over https only`,
    );
  }
  if ((url.pathname !== "/" && url.pathname !== "") || url.search || url.hash) {
    throw new Error(`${what} '${value}' carries a path - give the origin only`);
  }
  return url.origin;
}

function httpsUrl(value, what) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${what} '${value}' is no URL`);
  }
  if (url.protocol !== "https:")
    throw new Error(`${what} '${value}' is not https`);
  return url;
}

function absolutePath(value, what) {
  if (
    !value.startsWith("/") ||
    value.startsWith("//") ||
    /[?#\s]/.test(value)
  ) {
    throw new Error(`${what} '${value}' is no absolute path (without query)`);
  }
  return value;
}

function limit(value, max, what) {
  if (!value || value.length > max) {
    throw new Error(`${what} must have 1 to ${max} characters`);
  }
  return value;
}

// A stable id for a stable target: the same host, page and app give the
// same add-in, so a manifest written again replaces the installed one
// instead of adding a second. Shaped as a version 5 UUID.
export function stableId(seed) {
  const h = createHash("sha1")
    .update(`abap2UI5/office-addin:${seed}`)
    .digest("hex")
    .split("");
  h[12] = "5";
  h[16] = ((parseInt(h[16], 16) & 0x3) | 0x8).toString(16);
  const s = h.join("");
  return `${s.slice(0, 8)}-${s.slice(8, 12)}-${s.slice(12, 16)}-${s.slice(16, 20)}-${s.slice(20, 32)}`;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// options (as parseArgs returns them) -> { xml, values }
export function buildManifest(opts) {
  if (!opts.host)
    throw new Error("--host is required (https://<your SAP system>)");
  const host = httpsOrigin(opts.host, "--host");
  const path = absolutePath(opts.path || DEFAULTS.path, "--path");
  const app = String(opts.app || DEFAULTS.app).toUpperCase();
  if (!CLASS_NAME.test(app))
    throw new Error(`--app '${app}' is no ABAP class name`);

  const query = new URLSearchParams();
  query.set("app", app);
  if (opts.endpoint)
    query.set("endpoint", absolutePath(opts.endpoint, "--endpoint"));
  if (opts.client) {
    if (!/^\d{3}$/.test(opts.client))
      throw new Error(`--client '${opts.client}' is no client`);
    query.set("sap-client", opts.client);
  }
  if (opts.language) {
    if (!/^[A-Za-z]{1,2}$/.test(opts.language)) {
      throw new Error(`--language '${opts.language}' is no language key`);
    }
    query.set("sap-language", opts.language);
  }
  if (opts.theme) {
    if (!/^[a-z][a-z0-9_]*$/.test(opts.theme))
      throw new Error(`--theme '${opts.theme}' is no theme`);
    query.set("sap-ui-theme", opts.theme);
  }
  for (const p of opts.param || []) {
    const eq = p.indexOf("=");
    const name = eq > 0 ? p.slice(0, eq) : "";
    if (
      !/^[A-Za-z][A-Za-z0-9_]*$/.test(name) ||
      /^(app|endpoint)$/i.test(name)
    ) {
      throw new Error(
        `--param '${p}' is no name=value (app and endpoint are taken)`,
      );
    }
    query.append(name, p.slice(eq + 1));
  }
  const pageUrl = new URL(path, host);
  const sourceUrl = `${pageUrl.href}?${query.toString()}`;
  const commandsUrl = new URL("commands.html", pageUrl).href;

  const domains = [host];
  for (const d of opts["app-domain"] || []) {
    const origin = httpsOrigin(d, "--app-domain");
    if (!domains.includes(origin)) domains.push(origin);
  }

  let icons = opts.icons || DEFAULTS.icons;
  if (!icons.endsWith("/")) icons += "/";
  httpsUrl(icons, "--icons");

  const id = opts.id || stableId(`${host}${path}?app=${app}`);
  if (!GUID.test(id)) throw new Error(`--id '${id}' is no GUID`);
  const version = opts.version || DEFAULTS.version;
  if (!VERSION.test(version))
    throw new Error(`--version '${version}' is not n.n.n.n`);
  const supportUrl = httpsUrl(
    opts["support-url"] || DEFAULTS.supportUrl,
    "--support-url",
  ).href;

  const values = {
    ID: id,
    VERSION: version,
    PROVIDER: limit(opts.provider || DEFAULTS.provider, SHORT, "--provider"),
    NAME: limit(opts.name || DEFAULTS.name, SHORT, "--name"),
    DESCRIPTION: limit(
      opts.description || DEFAULTS.description,
      LONG,
      "--description",
    ),
    BUTTON: limit(opts.button || DEFAULTS.button, SHORT, "--button"),
    SUPPORT_URL: supportUrl,
    SOURCE_URL: sourceUrl,
    COMMANDS_URL: commandsUrl,
  };
  for (const size of ICON_SIZES)
    values[`ICON_${size}`] = `${icons}icon-${size}.png`;

  let xml = readFileSync(TEMPLATE, "utf8");
  // the template's own comment is for the template; the result says where
  // it came from instead
  xml = xml.replace(
    /<!--[\s\S]*?-->\n/,
    "<!-- Written by abap2UI5/office-addin scripts/manifest.mjs - " +
      "change the options and write it again rather than editing it. -->\n",
  );
  xml = xml.replace(
    "{{APP_DOMAINS}}",
    domains.map((d) => `    <AppDomain>${escapeXml(d)}</AppDomain>`).join("\n"),
  );
  xml = xml.replace(/\{\{([A-Z0-9_]+)\}\}/g, (m, name) => {
    if (!(name in values))
      throw new Error(`template placeholder ${m} has no value`);
    return escapeXml(values[name]);
  });
  return { xml, values, domains };
}

function main() {
  let opts;
  try {
    opts = parseArgs(process.argv.slice(2));
    const { xml, values } = buildManifest(opts);
    const out = opts.out || DEFAULTS.out;
    if (out === "-") {
      process.stdout.write(xml);
    } else {
      writeFileSync(out, xml);
      console.error(
        `manifest: ${out} - task pane ${values.SOURCE_URL}, id ${values.ID}`,
      );
    }
  } catch (e) {
    console.error(`manifest: ${e.message}`);
    process.exit(1);
  }
}

if (
  process.argv[1] &&
  resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  main();
}
