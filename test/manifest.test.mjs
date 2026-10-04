// The manifest generator (scripts/manifest.mjs) and what it writes, without
// a network: the XML parses, Microsoft's own manifest reader
// (office-addin-manifest, the package of `office-addin-manifest validate`)
// reads the fields Office needs, the elements stand in the order the schema
// prescribes, every resid resolves, and the options are checked. The
// validation service itself (online) runs in CI - .github/workflows/ci.yaml,
// job manifest.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createRequire } from "node:module";
import {
  buildManifest,
  parseArgs,
  stableId,
  DEFAULTS,
} from "../scripts/manifest.mjs";

const require = createRequire(import.meta.url);
const { OfficeAddinManifest } = require("office-addin-manifest");
const xml2js = require("xml2js");

const HOST = "https://my.sap.example";
const build = (args) => buildManifest(parseArgs(args));

async function parse(xml) {
  return xml2js.parseStringPromise(xml, {
    explicitChildren: true,
    preserveChildrenOrder: true,
  });
}

test("writes a manifest Microsoft's reader understands", async () => {
  const { xml, values } = build([
    "--host",
    HOST,
    "--path",
    "/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html",
    "--app",
    "z2ui5_cl_xl_demo",
    "--client",
    "100",
  ]);
  assert.doesNotMatch(xml, /\{\{/);
  const dir = mkdtempSync(join(tmpdir(), "z2ui5xl-manifest-"));
  try {
    const file = join(dir, "manifest.xml");
    writeFileSync(file, xml);
    const info = await OfficeAddinManifest.readManifestFile(file);
    assert.equal(info.id, values.ID);
    assert.equal(info.officeAppType, "TaskPaneApp");
    assert.deepEqual(info.hosts, ["Workbook"]);
    assert.deepEqual(info.appDomains, [HOST]);
    assert.equal(info.permissions, "ReadWriteDocument");
    assert.equal(info.version, "1.0.0.0");
    assert.equal(
      info.defaultSettings.sourceLocation,
      `${HOST}/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html?app=Z2UI5_CL_XL_DEMO&sap-client=100`,
    );
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test("the elements of OfficeApp stand in the schema's order", async () => {
  const { xml } = build(["--host", HOST]);
  const doc = await parse(xml);
  const names = doc.OfficeApp.$$.map((c) => c["#name"]);
  // "Basic task pane add-in element ordering" (Office Add-ins docs,
  // develop/manifest-element-ordering.md) - the ones this manifest has
  const order = [
    "Id",
    "Version",
    "ProviderName",
    "DefaultLocale",
    "DisplayName",
    "Description",
    "IconUrl",
    "HighResolutionIconUrl",
    "SupportUrl",
    "AppDomains",
    "Hosts",
    "Requirements",
    "DefaultSettings",
    "Permissions",
    "VersionOverrides",
  ];
  assert.deepEqual(names, order);
  const form =
    doc.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0];
  assert.deepEqual(
    form.$$.map((c) => c["#name"]),
    ["GetStarted", "FunctionFile", "ExtensionPoint"],
  );
});

test("every resid has its resource, every URL is https", async () => {
  const { xml } = build([
    "--host",
    HOST,
    "--app-domain",
    "https://idp.example",
  ]);
  const ids = new Set(
    [...xml.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map(
      (m) => m[1],
    ),
  );
  const resids = [...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]);
  assert.ok(resids.length > 5);
  for (const r of resids) assert.ok(ids.has(r), `resid ${r} has no resource`);
  for (const m of xml.matchAll(/DefaultValue="(http[^"]*)"/g)) {
    assert.match(m[1], /^https:\/\//, m[1]);
  }
  assert.match(xml, /<AppDomain>https:\/\/idp\.example<\/AppDomain>/);
  // the task pane button opens the page SourceLocation names
  const source = /<SourceLocation DefaultValue="([^"]+)"/.exec(xml)[1];
  assert.match(
    xml,
    new RegExp(
      `<bt:Url id="Taskpane.Url" DefaultValue="${source.replace(/[?.]/g, "\\$&")}"`,
    ),
  );
  assert.match(
    xml,
    /<bt:Url id="Commands.Url" DefaultValue="https:\/\/my\.sap\.example\/sap\/bc\/ui5_ui5\/sap\/z2ui5_xl\/commands\.html"/,
  );
});

test("the id is stable for a target and differs between targets", () => {
  const a = build(["--host", HOST]).values.ID;
  assert.equal(build(["--host", HOST]).values.ID, a);
  assert.notEqual(
    build(["--host", HOST, "--app", "Z2UI5_CL_OTHER"]).values.ID,
    a,
  );
  assert.match(
    a,
    /^[0-9a-f]{8}-[0-9a-f]{4}-5[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/,
  );
  assert.equal(stableId("x"), stableId("x"));
  const given = "0b6b3f0e-4b8e-4d2a-9d55-2f0a8f3c1e11";
  assert.equal(build(["--host", HOST, "--id", given]).values.ID, given);
});

test("parameters reach the task pane URL, escaped", () => {
  const { values, xml } = build([
    "--host",
    HOST,
    "--theme",
    "sap_fiori_3",
    "--language",
    "DE",
    "--endpoint",
    "/sap/bc/z2ui5_custom",
    "--param",
    "plant=1000",
    "--param",
    "note=a&b <c>",
    "--name",
    'Ein "Name"',
  ]);
  const url = new URL(values.SOURCE_URL);
  assert.equal(url.searchParams.get("sap-ui-theme"), "sap_fiori_3");
  assert.equal(url.searchParams.get("sap-language"), "DE");
  assert.equal(url.searchParams.get("endpoint"), "/sap/bc/z2ui5_custom");
  assert.equal(url.searchParams.get("plant"), "1000");
  assert.equal(url.searchParams.get("note"), "a&b <c>");
  assert.match(xml, /DisplayName DefaultValue="Ein &quot;Name&quot;"/);
  assert.doesNotMatch(xml, /<c>/);
});

test("refuses what Office or the host page would refuse", () => {
  const bad = [
    [[], /--host is required/],
    [["--host", "http://my.sap.example"], /not https/],
    [["--host", "https://my.sap.example/sap"], /origin only/],
    [["--host", HOST, "--app", "zcl bad"], /no ABAP class name/],
    [["--host", HOST, "--path", "sap/x.html"], /no absolute path/],
    [["--host", HOST, "--path", "//evil.example/x.html"], /no absolute path/],
    [["--host", HOST, "--client", "1"], /no client/],
    [["--host", HOST, "--id", "nope"], /no GUID/],
    [["--host", HOST, "--version", "1.0.0.0.0"], /not n.n.n.n/],
    [["--host", HOST, "--param", "app=X"], /are taken/],
    [["--host", HOST, "--app-domain", "http://idp.example"], /not https/],
    [["--host", HOST, "--icons", "http://cdn.example/"], /not https/],
    [["--host", HOST, "--name", "x".repeat(126)], /1 to 125/],
    [["--host", HOST, "--bogus", "1"], /unknown option/],
  ];
  for (const [args, message] of bad) {
    assert.throws(() => build(args), message, args.join(" "));
  }
});

test("defaults are the BSP of this repository and the demo app", () => {
  const { values } = build(["--host", HOST]);
  assert.equal(
    values.SOURCE_URL,
    `${HOST}${DEFAULTS.path}?app=${DEFAULTS.app}`,
  );
  assert.match(values.ICON_32, /\/manifest\/icons\/icon-32\.png$/);
});
