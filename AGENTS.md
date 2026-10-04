# AGENTS.md — office-addin

Single source of truth for agents working on **abap2UI5 for Excel**: an
Office add-in (task pane) that runs abap2UI5 apps in Microsoft Excel with the
real abap2UI5 UI5 frontend, and lets them read and write the workbook
through the frontend's custom control `z2ui5.cc.ExcelBridge`. `CLAUDE.md`
next to this file is a pointer at it, nothing more. README.md is the
documentation for admins and app developers; keep it in step with every
change.

**Language:** English for all code, comments, docs, commit messages, PRs.

## What this repository is

A source repository (abap2UI5 CONVENTIONS §1), pulled into SAP systems with
abapGit. It holds:

- the **host page** - the task pane's document: Office.js, the system's UI5
  and the abap2UI5 app embedded with `z2ui5.embed.Container`
  (`@abap2ui5/embed-control`) - as the BSP application `Z2UI5_XL`;
- the **demo app** `Z2UI5_CL_XL_DEMO`;
- the **manifest** template and its generator (`npm run manifest`);
- the tests: Node tests and a Playwright smoke test against a stand-in for
  Office.js.

Three things are **not** here and must not be copied in:

- **The control.** `z2ui5.cc.ExcelBridge` is abap2UI5's
  (`app/webapp/cc/ExcelBridge.js`, pull request #2847). A change to what it
  does is a pull request there (merged into abap2UI5 `main` as bcce7d6, in
  no release yet); do not work around it in the host page.
  `test/e2e/vendor/ExcelBridge.js` is a byte copy for the smoke test only
  (header: source commit), never shipped, never edited - delete it with the
  bump of `@abap2ui5/node-runtime` to the first release that has the control.
- **The abap2UI5 frontend.** The embed control loads it from the system
  (`<endpoint>?z2ui5-bundle`); nothing here pins or vendors it.
- **The embed control's source.** `src/02` carries a copy of
  `@abap2ui5/embed-control` because a BSP has to - generated from
  `node_modules`, in the version `package-lock.json` names, never edited.

## Layout

| Path | |
|---|---|
| `webapp/` | The host page, the source of the BSP: `index.html`, `host.js` (the UI5 module that starts the app), `office-init.js` and `history-cache.js` (around Office.js), `host.css`, `commands.html` (the manifest's FunctionFile) |
| `src/01/` | `Z2UI5_CL_XL_DEMO`, the demo app |
| `src/02/` | **Generated**: the BSP `Z2UI5_XL` (pages, page directory, `UI5RepositoryPathMapping.xml`), its two ICF nodes, `package.devc.xml` - `npm run bsp` |
| `manifest/` | `manifest.template.xml` and the icons it points at (`icons/icon-<size>.png`) |
| `scripts/build-bsp.mjs` | `webapp/` + the embed control -> `src/02` (`npm run bsp`, `bsp:check`), or -> `dist/` (`npm run dist`) |
| `scripts/manifest.mjs` | The manifest generator (`npm run manifest -- --host ...`) |
| `test/*.test.mjs` | Node tests: the manifest, the BSP format, the host page rules (`npm test`) |
| `test/e2e/` | The Playwright smoke test: `server.mjs` (one origin like the system: the BSP as delivered, `@abap2ui5/node-runtime` with the demo, UI5 from npm), `office-stub.js`, `addin.spec.mjs`, `build-backend.mjs`, `vendor/` |
| `examples/approuter/` | Serving the host page without a BSP - documented, not tested |
| `abaplint.jsonc`, `.github/abaplint/` | ABAP Standard (v750), ABAP Cloud, 7.02 |
| `abap2ui5lint.jsonc` | The abap2UI5 linter over `src/01` |
| `.github/workflows/` | `check-abap.yaml`, `check-addin.yaml`, `test-e2e.yaml` (see "Gates") |

## Rules for the host page (`webapp/`)

These break silently - in Excel, where no test of ours runs. `test/bsp.test.mjs`
holds most of them; keep it in step.

- **Office.js in the head, from Microsoft's CDN, before UI5**:
  `history-cache.js`, then
  `https://officeapis.public.onecdn.static.microsoft/1/office.js`, then
  `office-init.js`, then the UI5 bootstrap. Office.js must be in the top
  document and initialized before the body; it nulls
  `history.pushState`/`replaceState`, which the two files keep and restore
  (Microsoft's documented workaround). `office-init.js` calls
  `Office.onReady( )` at once - some hosts show no task pane until it is
  called.
- **No inline script, no inline handler, no `eval`.** The page must run under
  a Content-Security-Policy without `'unsafe-inline'`; the smoke test serves
  it with the README's policy and fails on a CSP violation.
- **`data-sap-ui-frameOptions="allow"`.** Excel on the web frames the page;
  what may frame it is the server's `frame-ancestors`, not UI5's.
- **UI5 1.71 is the floor** (abap2UI5's and the embed control's): no module,
  API or browser feature newer than what UI5 1.71 and the Office webviews
  (WebView2, WKWebView) have. The smoke test runs on 1.71 and 1.136.
- **Same origin.** The page starts the app through the embed control with a
  path on its own origin; never accept a URL from outside for the endpoint or
  for a script. The `app` parameter is checked against an ABAP class name
  before anything starts.
- **The page decides nothing about the app.** It picks the class, hands over
  the URL parameters and `office_host` / `office_platform`, and shows one
  plain-DOM message line (textContent only) when Office or the app is not
  there. Everything else is the ABAP app's.
- **Long lines break the BSP.** A source line over 255 characters is refused
  by `build-bsp.mjs` (the system keeps pages as 255-character rows).
- After any change to `webapp/` or a bump of `@abap2ui5/embed-control`:
  `npm run bsp` and commit `src/02` with it.

## Rules for the BSP (`src/02`)

- **Generated, never edited.** `npm run bsp:check` (CI) fails on any
  difference. The format is abapGit's serialization - the one abap2UI5's
  `tools/app2bsp` writes: pages padded to 255 characters, LF, no final
  newline, XML with a UTF-8 BOM, ICF file names `<name padded to 15><first 25
  hex of sha1(URL)>.sicf.xml`.
- **Text pages only.** abapGit's WAPA serializer stores pages as text, so no
  image or other binary goes into `webapp/` - the icons are served elsewhere
  (`--icons`).
- **The BSP's identity is `Z2UI5_XL`** (BSP, ICF nodes, the default `--path`
  of the manifest, `BSP_PATH` of the test server). Renaming it touches all
  of them and the README.
- `/sap/bc/ui5_ui5/sap/z2ui5_xl/resources/` is the system's UI5 - the page
  bootstraps from `resources/sap-ui-core.js`, and the embed control lives
  under `thirdparty/`, never under `resources/`.

## Rules for the demo app (`src/01`)

- Built like any abap2UI5 app (abap2UI5 `docs/agents/building-apps.md`):
  `z2ui5_cl_ui5_view_builder` in the house chain layout (`npm run fmt:chains`),
  one IF chain of lifecycle checks, `follow_up_action( )` with
  `cs_event-control_by_id` for `write` / `read`.
- **Narrow width.** The task pane is about 350 px wide: no `OverflowToolbar`
  for the main actions (they vanish into the overflow menu), popin columns
  for the tables.
- Lint-clean for ABAP Standard, ABAP Cloud and the 7.02 downport, and for the
  abap2UI5 linter at `failOn: warning`. Class names `^Z2UI5_CL_XL_`, within
  25 characters (the budget of abap2UI5's namespace rename).
- `z2ui5.cc.ExcelBridge` is not mirrored in the released linter
  (`lib/cc-controls.mjs` of abap2UI5/linter): its attributes get no verdict
  (no finding, no waiver needed). When the linter mirrors it, they are
  checked - fix what it finds rather than waiving it.
- An enum-typed property never gets an initial value from a binding
  (`MessageStrip type` = `""` terminates the app) - seed it in `model_init`.

## Rules for the manifest

- **Add-in only manifest (XML).** It installs on every Excel platform; the
  unified manifest does not install on perpetual Office for Windows. A JSON
  variant would be an addition, never a replacement.
- The element order is the schema's (Microsoft: *manifest element ordering*);
  `test/manifest.test.mjs` checks it - a new element goes into its place and
  into the test.
- Every URL is https; the id is derived from host, path and app unless given,
  so writing a manifest again replaces the installed add-in.
- The permission stays `ReadWriteDocument`; the ExcelApi requirement is the
  newest API set the bridge uses (1.4 today) - raise it when the control
  starts to use something newer.

## Gates

```bash
npm ci
npm run check        # lint, format:check, test, bsp:check, abaplint, abaplint:cloud, check:abap2ui5
npm run e2e:build && npm run e2e
```

| Workflow (`name:`) | Does |
|---|---|
| `check-abap.yaml` (`check-abap`) | abaplint Standard, Cloud and 7.02 (downported in CI), the abap2UI5 linter |
| `check-addin.yaml` (`check-addin`) | ESLint, Prettier, `npm test`, `bsp:check`; a generated manifest through Microsoft's validation service |
| `test-e2e.yaml` (`test-e2e`) | The Playwright smoke test, UI5 1.136 and 1.71 |

All three run on pull requests, pushes to `main` and weekly (the lint
configs follow abap2UI5 `main` and `702`).

`npm run check` leaves out, as CONVENTIONS §3 asks to name:

- **the 7.02 lint** - `npm run downport` rewrites `src/` in place; run it on a
  throwaway copy;
- **the online manifest validation** - it posts the manifest to Microsoft's
  service (`office-addin-manifest validate`); `npm test` checks the manifest
  offline;
- **the e2e test** - it needs a Chromium and the transpiled backend
  (`npm run e2e:build`, which fetches open-abap-core).

Pins, named as CONVENTIONS §4 asks: `@abap2ui5/node-runtime` and
`@abaplint/transpiler-cli` move together (the transpiler must be the version
the runtime's `package.json` names; `build-backend.mjs` refuses others);
`@openui5/sap.m` 1.136 and `ui5-1.71-sap.m` (an npm alias of
`@openui5/sap.m@1.71`) are the two UI5 versions of the smoke test;
`@abap2ui5/embed-control` is exact, because its copy is committed in `src/02`.

**What the gates cannot prove**: that a real Excel loads the add-in, the real
Office.js under the page's CSP, logon and cookies in Excel on the web. The
README's "What is tested" lists it; keep that list honest when something
gets verified by hand.

## Conventions

- All text files are LF-only (`.gitattributes`); JavaScript, JSON and CSS are
  formatted with Prettier (`.prettierrc`); Markdown is not.
- Node 22 (`.nvmrc`), scripts are ESM `.mjs` in `scripts/`.
- Actions are pinned to a commit SHA with the version in a trailing comment;
  Dependabot moves them. Workflows ask for `contents: read` only.
- `changelog.txt`: `unreleased` on top, one line per user-visible change.
- The ecosystem-wide rules live in abap2UI5's
  [CONVENTIONS.md](https://github.com/abap2UI5/abap2UI5/blob/main/.github/shared/CONVENTIONS.md).
