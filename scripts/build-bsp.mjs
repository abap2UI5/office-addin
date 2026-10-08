#!/usr/bin/env node
// Builds the BSP application Z2UI5_XL - the add-in's host page, served by
// the SAP system itself - into src/02 in abapGit's file format:
//
//   webapp/*                         the host page, as git has it
//   thirdparty/z2ui5/embed/          z2ui5.embed.Container and its stylesheet,
//                                    from node_modules/@abap2ui5/embed-control
//                                    (the version package-lock.json names) -
//                                    where `ui5 build` puts it in any app that
//                                    takes the package
//   UI5RepositoryPathMapping.xml     the page list of the UI5 repository
//   z2ui5_xl.wapa.xml                the BSP and its page directory
//   two .sicf.xml                    /sap/bc/ui5_ui5/sap/z2ui5_xl/ and
//                                    /sap/bc/bsp/sap/z2ui5_xl/
//   package.devc.xml
//
//   npm run bsp          write src/02
//   npm run bsp:check    build into a temporary folder and fail when src/02
//                        differs - the CI gate; src/02 is generated, never
//                        edited by hand
//   npm run dist         the same pages as plain files in dist/, for an
//                        approuter or a reverse proxy instead of a BSP
//
// The file format is abapGit's serialization of a BSP, the one abap2UI5's
// tools/app2bsp writes for abap2UI5/frontend: every page line padded to 255
// characters, LF, no newline after the last line, the XML files with a
// UTF-8 BOM - so a pull into a system and a re-serialization produce no diff.
// A source line longer than 255 characters is refused instead of wrapped:
// the system keeps a page as 255-character rows, and a wrapped line of
// JavaScript could come back as two.

import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

export const BSP = "z2ui5_xl";
const BSP_UP = BSP.toUpperCase();
const TEXT = "abap2UI5 Excel add-in - host page (generated)";
const DEVC_TEXT = "abap2UI5 Excel add-in - host page";
const ICF_TEXT = "abap2UI5 Excel add-in";
const START_PAGE = "index.html";
const LINE = 255;
const BOM = "\uFEFF";
// CL_O2_API_PAGES=>CREATE_NEW_PAGE refuses any other page name with
// sy-subrc=2 (invalid_name), and abapGit then fails the whole WAPA import -
// a hyphen is enough (history-cache.js did it)
export const VALID_PAGE_NAME = /^[A-Za-z0-9_./]+$/;

// the place of the control in the app - its ui5.yaml serves it there, and
// index.html registers it there (resource root z2ui5.embed)
const THIRDPARTY = "thirdparty/z2ui5/embed";
const CONTROL_FILES = ["Container.js", "Container.css"];

function listFiles(dir, base = dir) {
  const files = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...listFiles(full, base));
    else if (entry.isFile())
      files.push(relative(base, full).split("\\").join("/"));
  }
  return files;
}

// page name -> source text
export function collectPages() {
  const pages = new Map();
  const webapp = join(root, "webapp");
  for (const name of listFiles(webapp).sort()) {
    pages.set(name, readFileSync(join(webapp, name), "utf8"));
  }
  const control = join(
    root,
    "node_modules",
    "@abap2ui5",
    "embed-control",
    "src",
  );
  if (!existsSync(control)) {
    throw new Error(
      "node_modules/@abap2ui5/embed-control is missing - run npm ci",
    );
  }
  for (const name of CONTROL_FILES) {
    pages.set(
      `${THIRDPARTY}/${name}`,
      readFileSync(join(control, name), "utf8"),
    );
  }
  return pages;
}

function pageFileName(pageName) {
  return `${BSP}.wapa.${pageName.replace(/\//g, "_-").toLowerCase()}`;
}

function toBspPage(name, content) {
  const lines = content.replace(/\r\n?/g, "\n").split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  return lines
    .map((line, i) => {
      if (line.length > LINE) {
        throw new Error(
          `${name}:${i + 1} has ${line.length} characters - a BSP page line ` +
            `takes ${LINE}; break the line in the source`,
        );
      }
      return line.padEnd(LINE);
    })
    .join("\n");
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function byUpperCase(a, b) {
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

function pageItem(name) {
  // the start page carries MIMETYPE / IS_START_PAGE instead of PAGETYPE -
  // abapGit's WAPA serialization
  const type =
    name === START_PAGE
      ? [
          "      <MIMETYPE>text/html</MIMETYPE>",
          "      <IS_START_PAGE>X</IS_START_PAGE>",
        ]
      : ["      <PAGETYPE>X</PAGETYPE>"];
  return [
    "    <item>",
    "     <ATTRIBUTES>",
    `      <APPLNAME>${BSP_UP}</APPLNAME>`,
    `      <PAGEKEY>${escapeXml(name.toUpperCase())}</PAGEKEY>`,
    `      <PAGENAME>${escapeXml(name)}</PAGENAME>`,
    ...type,
    "      <LAYOUTLANGU>E</LAYOUTLANGU>",
    "      <VERSION>A</VERSION>",
    "      <LANGU>E</LANGU>",
    "     </ATTRIBUTES>",
    "    </item>",
  ];
}

function wapaXml(pageNames) {
  return (
    BOM +
    [
      '<?xml version="1.0" encoding="utf-8"?>',
      '<abapGit version="v1.0.0" serializer="LCL_OBJECT_WAPA" serializer_version="v1.0.0">',
      ' <asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">',
      "  <asx:values>",
      "   <ATTRIBUTES>",
      `    <APPLNAME>${BSP_UP}</APPLNAME>`,
      "    <APPLCLAS>/UI5/CL_UI5_BSP_APPLICATION</APPLCLAS>",
      `    <APPLEXT>${BSP_UP}</APPLEXT>`,
      "    <SECURITY>X</SECURITY>",
      "    <ORIGLANG>E</ORIGLANG>",
      "    <MODIFLANG>E</MODIFLANG>",
      `    <TEXT>${escapeXml(TEXT)}</TEXT>`,
      "   </ATTRIBUTES>",
      "   <PAGES>",
      ...[...pageNames].sort(byUpperCase).flatMap(pageItem),
      "   </PAGES>",
      "  </asx:values>",
      " </asx:abap>",
      "</abapGit>",
      "",
    ].join("\n")
  );
}

// Every folder and file of the app, as the UI5 repository upload
// (/UI5/UI5_REPOSITORY_LOAD) records them.
function pathMapping(pageNames) {
  const entries = new Map();
  for (const name of pageNames) {
    const parts = name.split("/");
    for (let i = 1; i < parts.length; i += 1) {
      entries.set(parts.slice(0, i).join("/"), true);
    }
    entries.set(name, false);
  }
  const lines = [
    '<?xml version="1.0"?>',
    "",
    '<UI5RepMapping version="1.0" xmlns="sap.ui5.tools.repository.mapping">',
    " <MappingEntries>",
  ];
  for (const path of [...entries.keys()].sort()) {
    const folder = entries.get(path);
    lines.push(
      "  <MappingEntry",
      `   path              = "${escapeXml(path)}"`,
      `   is_folder         = "${folder ? "X" : ""}"`,
      `   internal_rep      = "${folder ? "" : "B"}"`,
      `   internal_rep_path = "${folder ? "" : escapeXml(path)}" />`,
    );
  }
  lines.push(" </MappingEntries>", "</UI5RepMapping>");
  return lines.join("\n");
}

// abapGit names an ICF node's file by the node name, padded to 15, and the
// first 25 hex characters of the SHA-1 of its URL
export function sicfFileName(url) {
  const hash = createHash("sha1").update(url).digest("hex").slice(0, 25);
  return `${BSP.padEnd(15)}${hash}.sicf.xml`;
}

function sicfXml(url) {
  return (
    BOM +
    [
      '<?xml version="1.0" encoding="utf-8"?>',
      '<abapGit version="v1.0.0" serializer="LCL_OBJECT_SICF" serializer_version="v1.0.0">',
      ' <asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">',
      "  <asx:values>",
      `   <URL>${url}</URL>`,
      "   <ICFSERVICE>",
      `    <ICF_NAME>${BSP_UP}</ICF_NAME>`,
      `    <ORIG_NAME>${BSP}</ORIG_NAME>`,
      "   </ICFSERVICE>",
      "   <ICFDOCU>",
      `    <ICF_NAME>${BSP_UP}</ICF_NAME>`,
      "    <ICF_LANGU>E</ICF_LANGU>",
      `    <ICF_DOCU>${escapeXml(ICF_TEXT)}</ICF_DOCU>`,
      "   </ICFDOCU>",
      "  </asx:values>",
      " </asx:abap>",
      "</abapGit>",
      "",
    ].join("\n")
  );
}

function devcXml(text) {
  return (
    BOM +
    [
      '<?xml version="1.0" encoding="utf-8"?>',
      '<abapGit version="v1.0.0" serializer="LCL_OBJECT_DEVC" serializer_version="v1.0.0">',
      ' <asx:abap xmlns:asx="http://www.sap.com/abapxml" version="1.0">',
      "  <asx:values>",
      "   <DEVC>",
      `    <CTEXT>${escapeXml(text)}</CTEXT>`,
      "   </DEVC>",
      "  </asx:values>",
      " </asx:abap>",
      "</abapGit>",
      "",
    ].join("\n")
  );
}

export const ICF_URLS = [
  `/sap/bc/ui5_ui5/sap/${BSP}/`,
  `/sap/bc/bsp/sap/${BSP}/`,
];

// file name -> content of src/02
export function buildFiles() {
  const pages = collectPages();
  const files = new Map();
  const seen = new Map();
  pages.set("UI5RepositoryPathMapping.xml", pathMapping([...pages.keys()]));
  for (const [name, content] of pages) {
    if (!VALID_PAGE_NAME.test(name)) {
      throw new Error(
        `'${name}' is no BSP page name - CL_O2_API_PAGES takes only ` +
          `letters, digits, "_", "." and "/"; rename the file`,
      );
    }
    const file = pageFileName(name);
    // lower case and / -> _- are not injective: refuse instead of losing one
    if (seen.has(file)) {
      throw new Error(`'${name}' and '${seen.get(file)}' are both ${file}`);
    }
    seen.set(file, name);
    files.set(file, toBspPage(name, content));
  }
  files.set(`${BSP}.wapa.xml`, wapaXml(pages.keys()));
  for (const url of ICF_URLS) files.set(sicfFileName(url), sicfXml(url));
  files.set("package.devc.xml", devcXml(DEVC_TEXT));
  return files;
}

function write(dir, files) {
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(dir, { recursive: true });
  for (const [name, content] of files) writeFileSync(join(dir, name), content);
}

function main() {
  const args = process.argv.slice(2);
  const check = args.includes("--check");
  const dist = args.includes("--dist");
  if (args.some((a) => a !== "--check" && a !== "--dist") || (check && dist)) {
    console.error("build-bsp: takes --check, --dist or nothing");
    process.exit(2);
  }
  if (dist) {
    // the host page as plain files, for a server other than the SAP system
    // that puts it on the system's origin - an approuter, a reverse proxy
    // (README, "Without a BSP")
    const out = join(root, "dist");
    rmSync(out, { recursive: true, force: true });
    const pages = collectPages();
    for (const [name, content] of pages) {
      mkdirSync(dirname(join(out, name)), { recursive: true });
      writeFileSync(join(out, name), content);
    }
    console.log(`build-bsp: ${pages.size} files in dist/`);
    return;
  }
  const target = join(root, "src", "02");
  const files = buildFiles();
  if (!check) {
    write(target, files);
    console.log(`build-bsp: ${files.size} files in src/02 (BSP ${BSP_UP})`);
    return;
  }
  const scratch = mkdtempSync(join(tmpdir(), "z2ui5xl-bsp-"));
  try {
    write(scratch, files);
    const want = new Set(files.keys());
    const have = existsSync(target) ? new Set(readdirSync(target)) : new Set();
    const problems = [];
    for (const name of want) {
      if (!have.has(name)) problems.push(`missing: src/02/${name}`);
      else if (readFileSync(join(target, name), "utf8") !== files.get(name)) {
        problems.push(`differs: src/02/${name}`);
      }
    }
    for (const name of have) {
      if (!want.has(name)) problems.push(`not generated: src/02/${name}`);
    }
    if (problems.length) {
      console.error(problems.join("\n"));
      console.error(
        "build-bsp: src/02 is stale - run npm run bsp and commit the result",
      );
      process.exit(1);
    }
    console.log(`build-bsp: src/02 is up to date (${files.size} files)`);
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
}

if (
  process.argv[1] &&
  resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  main();
}
