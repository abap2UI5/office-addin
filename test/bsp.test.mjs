// The BSP of src/02 and the host page in it: what abapGit needs to pull it
// without a diff, and the rules of the host page that break silently - in
// Excel, where no test of ours runs (AGENTS.md, "Rules for the host page").
import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  buildFiles,
  sicfFileName,
  ICF_URLS,
  BSP,
} from "../scripts/build-bsp.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const webapp = (name) => readFileSync(join(root, "webapp", name), "utf8");

test("src/02 is what scripts/build-bsp.mjs builds", () => {
  const files = buildFiles();
  const dir = join(root, "src", "02");
  assert.deepEqual(readdirSync(dir).sort(), [...files.keys()].sort());
  for (const [name, content] of files) {
    assert.equal(
      readFileSync(join(dir, name), "utf8"),
      content,
      `src/02/${name} is stale - npm run bsp`,
    );
  }
});

test("pages in abapGit's format: 255-character rows, LF, no final newline", () => {
  for (const [name, content] of buildFiles()) {
    if (!name.startsWith(`${BSP}.wapa.`) || name.endsWith(".wapa.xml"))
      continue;
    assert.ok(!content.includes("\r"), name);
    assert.ok(!content.endsWith("\n"), name);
    for (const line of content.split("\n"))
      assert.equal(line.length, 255, name);
  }
});

test("every page is in the page directory, the start page is index.html", () => {
  const files = buildFiles();
  const wapa = files.get(`${BSP}.wapa.xml`);
  assert.ok(wapa.startsWith("﻿<?xml"));
  const pages = [...wapa.matchAll(/<PAGENAME>([^<]+)<\/PAGENAME>/g)].map(
    (m) => m[1],
  );
  for (const page of pages) {
    const file = `${BSP}.wapa.${page.replace(/\//g, "_-").toLowerCase()}`;
    assert.ok(files.has(file), `${page} -> ${file}`);
  }
  const pageFiles = [...files.keys()].filter(
    (f) => f.startsWith(`${BSP}.wapa.`) && f !== `${BSP}.wapa.xml`,
  );
  assert.equal(pageFiles.length, pages.length);
  assert.match(
    wapa,
    /<PAGENAME>index\.html<\/PAGENAME>\n\s+<MIMETYPE>text\/html<\/MIMETYPE>\n\s+<IS_START_PAGE>X<\/IS_START_PAGE>/,
  );
  assert.ok(pages.includes("thirdparty/z2ui5/embed/Container.js"));
});

test("the ICF nodes carry abapGit's file names", () => {
  assert.deepEqual(ICF_URLS, [
    "/sap/bc/ui5_ui5/sap/z2ui5_xl/",
    "/sap/bc/bsp/sap/z2ui5_xl/",
  ]);
  for (const url of ICF_URLS) {
    const hash = createHash("sha1").update(url).digest("hex").slice(0, 25);
    assert.equal(sicfFileName(url), `z2ui5_xl       ${hash}.sicf.xml`);
  }
});

test("the host page loads Office.js first, from Microsoft's CDN, in the head", () => {
  const html = webapp("index.html");
  const head = html.slice(0, html.indexOf("</head>"));
  const office = head.indexOf(
    'src="https://officeapis.public.onecdn.static.microsoft/1/office.js"',
  );
  const cache = head.indexOf('src="history-cache.js"');
  const init = head.indexOf('src="office-init.js"');
  const ui5 = head.indexOf('src="resources/sap-ui-core.js"');
  assert.ok(
    cache >= 0 && office > cache && init > office && ui5 > init,
    "history-cache, office.js, office-init, UI5",
  );
});

test("the host page has no inline script and no inline handler", () => {
  for (const name of ["index.html", "commands.html"]) {
    const html = webapp(name);
    for (const m of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)) {
      assert.match(m[1], /\bsrc=/, `${name}: a script without src`);
      assert.equal(m[2].trim(), "", `${name}: inline script`);
    }
    assert.doesNotMatch(
      html,
      /\son[a-z]+\s*=/i,
      `${name}: inline event handler`,
    );
  }
});

test("UI5 is bootstrapped from the system, frameable, with the add-in's roots", () => {
  const html = webapp("index.html");
  assert.match(html, /data-sap-ui-frameOptions="allow"/);
  assert.doesNotMatch(html, /frameOptions="(trusted|deny)"/);
  assert.match(html, /"z2ui5\.embed": "\.\/thirdparty\/z2ui5\/embed\/"/);
  assert.match(html, /data-sap-ui-oninit="module:z2ui5xl\/host"/);
});

test("host.js stays on the UI5 1.71 floor and away from eval", () => {
  const js =
    webapp("host.js") + webapp("office-init.js") + webapp("history-cache.js");
  assert.doesNotMatch(
    js,
    /\beval\s*\(|new Function|innerHTML|insertAdjacentHTML|document\.write/,
  );
  assert.doesNotMatch(js, /sap\/ui\/core\/(Lib|Element)"|sap\.ui\.core\.Lib\b/);
  assert.match(
    webapp("host.js"),
    /sap\.ui\.define\(\["z2ui5\/embed\/Container"\]/,
  );
});
