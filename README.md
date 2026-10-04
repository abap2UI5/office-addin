# abap2UI5 for Excel

**abap2UI5 apps in Microsoft Excel**: an Office add-in that shows an
[abap2UI5](https://github.com/abap2UI5/abap2UI5) app in Excel's task pane -
the real abap2UI5 UI5 frontend, served by your SAP system - and lets the app
read and write the workbook through the frontend's custom control
`z2ui5.cc.ExcelBridge`. The app stays an ABAP class; the add-in is a
manifest, one small host page and nothing else.

This is a **source repository** (abap2UI5 CONVENTIONS §1): you edit here, CI
checks every change. It holds the host page as a BSP application to pull
with abapGit, a demo app, the manifest generator and the tests.

> **Status: proof of concept.** CI proves the host page, the embedded app and
> the bridge against a stand-in for Office.js. Nobody has run it in a real
> Excel yet - see [What is tested](#what-is-tested).

- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Install (admins)](#install-admins)
- [The manifest](#the-manifest)
- [Deploy to users](#deploy-to-users)
- [Excel on the web](#excel-on-the-web)
- [Content-Security-Policy](#content-security-policy)
- [Without a BSP](#without-a-bsp)
- [Apps that use Excel (developers)](#apps-that-use-excel-developers)
- [Limits](#limits)
- [What is tested](#what-is-tested)
- [Development](#development)

## How it works

```
Excel ── task pane (WebView2 / WKWebView / an iframe of Excel on the web)
          │
          │ https://<sap-host>/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html?app=Z2UI5_CL_XL_DEMO
          ▼
        host page (BSP Z2UI5_XL, this repository)
          ├─ Office.js            from Microsoft's CDN, in the page's head
          ├─ UI5                  the system's own (resources/)
          └─ z2ui5.embed.Container  (@abap2ui5/embed-control)
                │  GET  /sap/bc/z2ui5?z2ui5-bundle   the abap2UI5 frontend
                │  POST /sap/bc/z2ui5                the roundtrips
                ▼
              the abap2UI5 app ── z2ui5.cc.ExcelBridge ── window.Excel (Office.js)
```

Office.js only works in the **top document** of the task pane, and
abap2UI5's own page (`/sap/bc/z2ui5`) neither loads foreign scripts nor lets
itself be framed (`frameOptions` trusted). So the add-in brings its own page:
it loads Office.js, boots UI5 with `frameOptions` allow, and embeds the
abap2UI5 component with the [embed control](https://github.com/abap2UI5/embed-control)
in **embedded mode** (the app leaves the URL alone). The frontend then runs in
the same window as Office.js, and `z2ui5.cc.ExcelBridge` finds `window.Excel`.

**The host page and the abap2UI5 service share one origin.** abap2UI5
refuses a POST from another origin (its CSRF check), and the embed control
refuses to load the frontend from another origin. That is why the page is a
BSP of the same system - or, without BSPs, served by an approuter or a
reverse proxy that puts both under one host ([Without a BSP](#without-a-bsp)).

## Requirements

- **abap2UI5 with `z2ui5.cc.ExcelBridge`** - abap2UI5 pull request #2847,
  on abap2UI5 `main` since commit bcce7d6. Until a release contains it (the
  first after 1.146.0 whose `changelog.txt` lists the control), install
  abap2UI5 from `main`. With an older
  abap2UI5 the demo app's view does not load (`z2ui5/cc/ExcelBridge` is
  missing). The embed control needs abap2UI5 1.145.0 or later
  (`?z2ui5-bundle`); this add-in needs the release with #2847 anyway.
- **UI5 1.71 or later** on the system (the floor of abap2UI5 and of the embed
  control). The host page asks for the theme `sap_horizon` (UI5 1.102+);
  below that, write the manifest with `--theme sap_fiori_3`.
- **HTTPS** with a certificate the users' machines trust - Office loads
  add-ins over https only.
- **Excel**: Microsoft 365, Excel 2019 or later on Windows or Mac, Excel on
  the web (best effort, see below). The manifest asks for ExcelApi 1.4 -
  the newest API set the bridge uses (`Range.getUsedRangeOrNullObject`) -
  which a volume-licensed Excel 2016 does not have.
  Windows uses WebView2 (Edge) for add-ins, Mac WKWebView (Microsoft:
  *Browsers and webview controls used by Office Add-ins*).
- For `npm run manifest`: Node.js 22 and `npm ci` in a clone of this
  repository.

## Install (admins)

1. **Pull this repository with abapGit** into a package of its own (for
   example `$Z2UI5_XL`, or a transportable `Z2UI5_XL`), on the system that
   runs abap2UI5. It brings
   - `Z2UI5_CL_XL_DEMO` - the demo app (`src/01`),
   - the BSP application `Z2UI5_XL` - the host page and the embed control
     (`src/02`, generated - see [Development](#development)),
   - the ICF nodes `/sap/bc/ui5_ui5/sap/z2ui5_xl` and `/sap/bc/bsp/sap/z2ui5_xl`.
2. **SICF**: check that `/sap/bc/ui5_ui5/sap/z2ui5_xl` is active, with the
   same logon settings as abap2UI5's `/sap/bc/z2ui5`.
3. **Try it in a browser** first:
   `https://<sap-host>/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html?app=Z2UI5_CL_XL_DEMO`
   shows the demo with a line "Not running in Excel". The page is the
   add-in's page - it works outside Office as a plain embedding, and the
   Excel buttons then report "Excel not available".
4. **Write the manifest** ([below](#the-manifest)) and
   [deploy it](#deploy-to-users).
5. For Excel on the web, set the [framing and cookie headers](#excel-on-the-web).

**Logon.** The task pane opens the page like a browser tab: whatever logon
the ICF node asks for happens in the task pane - on the desktop that is a
top-level window, so an SSO setup (Kerberos/SPNEGO, X.509, SAML) works as in
a browser. A SAML identity provider on another host has to be in the
manifest's `AppDomains` (`--app-domain https://idp.example`), or the desktop
opens it in the system browser instead of the task pane. The session cookie
carries the client: a manifest written with `--client` hands `sap-client` to
the host page, and the roundtrips follow it.

**ABAP Cloud** (SAP BTP, ABAP environment; S/4HANA Cloud public edition) has
no BSP applications: pull only `src/01` there and serve the host page
[without a BSP](#without-a-bsp).

## The manifest

Office add-ins are registered by a manifest. This repository writes the
**add-in only manifest** (XML), which every Excel platform installs - the
unified manifest for Microsoft 365 (JSON) is not supported by the perpetual
versions of Office on Windows (Office 2019, 2021, LTSC) and needs Microsoft
365 builds from 2501 on Windows and 16.103 on the Mac (Microsoft: *Office Add-ins with the unified app
manifest*, "Client and platform support"), which many SAP customers do not
have everywhere.

```bash
npm ci
npm run manifest -- --host https://my.sap.example \
  --path /sap/bc/ui5_ui5/sap/z2ui5_xl/index.html \
  --app Z2UI5_CL_XL_DEMO --client 100 --out manifest.xml
```

| Option | |
|---|---|
| `--host` | The system as the users' browser reaches it: https, origin only (required) |
| `--path` | The host page (default `/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html`) |
| `--app` | The abap2UI5 app the task pane starts (default `Z2UI5_CL_XL_DEMO`) |
| `--client`, `--language` | `sap-client`, `sap-language` for the page |
| `--theme` | `sap-ui-theme` - `sap_fiori_3` on a system below UI5 1.102 |
| `--endpoint` | The abap2UI5 service when it is not `/sap/bc/z2ui5` |
| `--param name=value` | A startup parameter for the app (`client->get( )-t_comp_params`); repeatable |
| `--app-domain` | Another origin the task pane may navigate to (identity provider); repeatable |
| `--icons` | Where the icons are served: a folder URL with `icon-16/32/64/80.png` of `manifest/icons/` (default: this repository on GitHub) |
| `--id` | The add-in id. Default: derived from host, path and app, so writing the manifest again keeps it |
| `--version` | `1.0.0.0` by default - raise it with every change to a deployed manifest |
| `--name`, `--description`, `--provider`, `--button`, `--support-url` | What users see |

One manifest per app: each `--app` gives an add-in of its own with its own
id and ribbon button. The template is `manifest/manifest.template.xml`
(one button "abap2UI5" on the Home tab, task pane, `ReadWriteDocument`).

**Icons.** Office fetches the icons from the manifest's URLs and caches them
(they must not be sent with `Cache-Control: no-store`). A BSP cannot carry
them through abapGit (abapGit stores BSP pages as text), so the default
points at `manifest/icons/` of this repository on GitHub. To keep them
inside your network, import the four PNGs as MIME objects into the BSP
`Z2UI5_XL` (SE80, folder `icons/`), or put them on any https server, and
write the manifest with `--icons https://my.sap.example/sap/bc/ui5_ui5/sap/z2ui5_xl/icons/`.

**Validation.** CI writes a manifest and runs Microsoft's
`office-addin-manifest validate` on it (an online service); `npm test`
checks it offline with Microsoft's manifest reader and the schema's element
order. To validate yours: `npx office-addin-manifest validate manifest.xml`.

## Deploy to users

Microsoft's documentation: *Deploy and publish Office Add-ins*.

- **Microsoft 365 admin center** (production): *Settings > Integrated apps >
  Upload custom apps*, app type *Office Add-in*, upload `manifest.xml`, pick
  users or groups. The add-in appears in their Excel without any client
  setup (it can take a few hours). Sovereign clouds use *Centralized
  Deployment* instead. A new manifest version with changed permissions
  needs the admin's consent again.
- **Sideload, Excel on Windows**: share a network folder, add it under
  *File > Options > Trust Center > Trust Center Settings > Trusted Add-in
  Catalogs* (tick *Show in Menu*), put `manifest.xml` into it, restart Excel,
  then *Home > Add-ins > Advanced > Shared Folder*.
- **Sideload, Excel on Mac**: copy `manifest.xml` to
  `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create the
  folder), restart Excel, *Home > Add-ins*.
- **Sideload, Excel on the web**: open a workbook, *Home > Add-ins > More
  Settings > Upload My Add-in*, choose `manifest.xml`.

After the add-in is installed, the **abap2UI5** button on the Home tab opens
the task pane.

## Excel on the web

Best effort. Excel on the web shows the task pane in a sandboxed iframe of
an Office page (`*.cloud.microsoft`, `*.officeapps.live.com`, ...), so two
things that do not matter on the desktop decide here:

1. **Framing.** The host page's response must allow those ancestors and must
   not carry `X-Frame-Options: DENY/SAMEORIGIN`:

   ```
   Content-Security-Policy: frame-ancestors 'self' https://*.cloud.microsoft https://*.officeapps.live.com https://*.office.com https://*.office365.com https://*.sharepoint.com
   ```

   The list is the set of Office domains known to host Excel on the web;
   **check it in your tenant** (browser devtools, the frame chain of the task
   pane) - Microsoft does not publish a fixed list for the add-in iframe's
   ancestors. The host page itself sets UI5's `frameOptions` to allow;
   abap2UI5's own node is not framed - the frontend comes as a script.
   Set the header on the system - see [CSP](#content-security-policy).
2. **Cookies.** In the iframe the SAP session cookies (`MYSAPSSO2`,
   `SAP_SESSIONID_<SID>_<client>`, `sap-usercontext`, `sap-contextid`) are
   **third-party cookies**. They need `SameSite=None; Secure`; check the SAP
   Notes on SameSite cookie handling for your kernel (ICM can add the
   attribute). Even then browsers may block them: Safari (ITP) does, Chrome
   and Edge partition storage and may phase out third-party cookies (Microsoft:
   *Develop your Office Add-in to work with ITP when using third-party
   cookies*). The logon itself is a page of the system in the iframe -
   identity providers that refuse to be framed fail there.

A sturdier logon for the web - the Office dialog API
(`Office.context.ui.displayDialogAsync`) opening the system's logon in a
top-level window, then a first-party session (Storage Access API, or a
token) - is an open follow-up, not built.

## Content-Security-Policy

The BSP sends no Content-Security-Policy of its own. Recommended for the host
page (`/sap/bc/ui5_ui5/sap/z2ui5_xl/`):

```
default-src 'self';
script-src 'self' https://officeapis.public.onecdn.static.microsoft https://appsforoffice.microsoft.com;
style-src 'self' 'unsafe-inline';
img-src 'self' data:;
font-src 'self' data:;
connect-src 'self';
frame-ancestors 'self' https://*.cloud.microsoft https://*.officeapps.live.com https://*.office.com https://*.office365.com https://*.sharepoint.com
```

- `officeapis.public.onecdn.static.microsoft` is Office.js's current CDN
  (`https://officeapis.public.onecdn.static.microsoft/1/office.js`, which the
  page loads); `appsforoffice.microsoft.com` is the legacy endpoint and still
  serves files Office.js may load. Microsoft: *Referencing the Office
  JavaScript API library* - it also says to allow Office.js's Trusted Types
  policy if your organisation enforces Trusted Types.
- The page has **no inline script and no inline handler** (CI checks it), so
  `script-src` needs no `'unsafe-inline'`; the abap2UI5 frontend is a script
  of the same origin and evaluates nothing from a string.
- `'unsafe-eval'`: UI5 1.71 needs it (the embed control says so for its
  hosts); a current UI5 does not.
- UI5 from a CDN instead of the system: add its origin to `script-src`,
  `style-src`, `font-src` and `img-src` and change the bootstrap `src` in
  `webapp/index.html`.
- Office.js may need more on some platforms (telemetry, locale files) -
  **not verified without a real Excel**; watch the browser console for CSP
  reports when you roll it out.

The CI smoke test serves the page with this policy (plus `'unsafe-eval'` for
UI5 from source files).

**Setting headers on the system.** One way is an ICM rewrite rule
(profile parameter `icm/HTTP/mod_<n>` with a rule file; the SAP Web
Dispatcher understands the same), for example:

```
# icm/HTTP/mod_0 = PREFIX=/sap/bc/ui5_ui5/sap/z2ui5_xl/, FILE=$(DIR_PROFILE)/z2ui5_xl_headers.txt
SetResponseHeader Content-Security-Policy "default-src 'self'; script-src 'self' https://officeapis.public.onecdn.static.microsoft https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self' data:; connect-src 'self'; frame-ancestors 'self' https://*.cloud.microsoft https://*.officeapps.live.com https://*.office.com https://*.office365.com https://*.sharepoint.com"
RemoveResponseHeader X-Frame-Options
```

Adapt it to your release (check the ICM documentation for the modification
file syntax) and test it in a browser before Excel - an example, not
something this repository could verify.

## Without a BSP

Where the system has no BSPs (SAP BTP, ABAP environment), or the landscape
puts an SAP approuter or a reverse proxy in front anyway, that server can
serve the host page **on the same origin** as the abap2UI5 service:

- `npm run dist` writes the host page with the embed control to `dist/`.
- `examples/approuter/` has an `xs-app.json` that serves `dist/`, routes
  `/sap/bc/z2ui5` to the system and `resources/` to UI5 - see its README.
- Write the manifest for that origin and path:
  `npm run manifest -- --host https://<approuter> --path /index.html`
  (plus `--endpoint` when the service has another path).

Documented, not tested by CI.

## Apps that use Excel (developers)

Any abap2UI5 app runs in the task pane - start it with `--app`. To talk to
the workbook, put `z2ui5.cc.ExcelBridge` into the view (it renders nothing)
and call its methods as frontend actions. From the demo
(`src/01/z2ui5_cl_xl_demo.clas.abap`):

```abap
DATA(view) = z2ui5_cl_ui5_view_builder=>factory(
    )->ele( n = `View` ns = `mvc`
        )->a( n = `xmlns`       v = `sap.m`
        )->a( n = `xmlns:mvc`   v = `sap.ui.core.mvc`
        )->a( n = `xmlns:z2ui5` v = `z2ui5.cc` ).

DATA(page) = view->ele( `Page`
    )->a( n = `title` v = `Materials` ).

page->tag( n = `ExcelBridge` ns = `z2ui5`
    )->a( n = `id`               v = `excel`
    )->a( n = `rows`             v = client->_bind( materials )
    )->a( n = `columns`          v = client->_bind( columns )
    )->a( n = `target`           v = `newSheet`
    )->a( n = `sheetName`        v = client->_bind( sheet_name )
    )->a( n = `selection`        v = client->_bind( selection )
    )->a( n = `selectionAddress` v = client->_bind( selection_address )
    )->a( n = `OnWritten`        v = client->_event( val = `XL_WRITTEN` arg = `${$parameters>/address}` )
    )->a( n = `OnRead`           v = client->_event( val = `XL_READ` arg = `${$parameters>/address}` )
    )->a( n = `OnError`          v = client->_event( val   = `XL_ERROR`
                                                     t_arg = VALUE #( ( `${$parameters>/message}` )
                                                                      ( `${$parameters>/code}` ) ) ) ).

" a button that writes without a roundtrip
page->tag( `Button`
    )->a( n = `text`  v = `Export to Excel`
    )->a( n = `press` v = client->follow_up_action( val   = client->cs_event-control_by_id
                                                    t_arg = VALUE #( ( `excel` )
                                                                     ( `write` ) ) ) ).
```

After a roundtrip, the same call as a statement:
`client->follow_up_action( val = client->cs_event-control_by_id t_arg = VALUE #( ( `excel` ) ( `read` ) ) ).`

| Property | |
|---|---|
| `rows` | The table to write (bind an ABAP table) |
| `columns` | Which fields, in which order, with which header: `MATNR,MAKTX`, or a table of `key` / `header` / `number_format` |
| `numberFormats` | Excel number formats along the columns (array) or by column key (object) |
| `target` | `A1` (default), any cell address, `selection`, or `newSheet` |
| `sheetName` | The sheet for a cell address, or the name of the new sheet |
| `asTable` | An Excel table with a header row (default true); `header` a header row without a table |
| `maxCells` | The cap per call (default 20000, at most 100000) |
| `selection` | `read( )` puts the selected cells here as rows `COL1`, `COL2`, ... - bind a table of a structure with components `col1`, `col2`, ...; `readText` reads the displayed text instead of values |
| `selectionAddress` | The address `read( )` read |
| `selectionChange` | `true`: fire `OnSelectionChange` (debounced) whenever the user selects something else |
| `available` | Turns true once Excel answered - bind a button's `visible` or `enabled` to it |

| Event | Parameters |
|---|---|
| `OnWritten` | `address`, `rowCount`, `columnCount` (header row included) |
| `OnRead` | `address`, `rowCount`, `columnCount` |
| `OnSelectionChange` | `address` |
| `OnError` | `message`, `code` (`NotAvailable` outside Excel, `TooLarge`, `NoData`, `InvalidTarget`, or Excel's own, e.g. a protected sheet) |

Values keep their types: strings are written as text (number format `@`, so
leading zeros stay and `=...` never becomes a formula), numbers as numbers,
ISO dates (an ABAP `d`) as Excel dates; a read returns numbers as numbers and
date cells as ISO dates. The control's own documentation is the header of
`app/webapp/cc/ExcelBridge.js` in abap2UI5.

The host page also hands the app `office_host` (`Excel`, or empty outside
Office) and `office_platform` (`PC`, `Mac`, `OfficeOnline`, ...) as startup
parameters, and every other URL parameter of the page except `app`,
`endpoint`, `sap-*` and Office's `_host_Info`.

The demo (`Z2UI5_CL_XL_DEMO`): a material list with *Export to Excel* (a new
sheet with an Excel table, per export), *Take selection* (the selected cells
into a second table), a *Follow selection* switch (enabled once Excel
answered) and a message strip for results and errors. It is lint-clean for
ABAP Standard, ABAP Cloud and the 7.02 downport.

## Limits

- **Size.** At most `maxCells` cells per write or read (20000 by default,
  100000 at most, header row included); more is refused with `OnError`
  `TooLarge`, nothing written. Excel on the web takes a few MB per request,
  and every row also travels through the abap2UI5 model.
- **Excel on the web** depends on third-party cookies and framing headers -
  best effort, see above.
- **One app per manifest**; the task pane shows the app the manifest names.
- **Not tested in a real Excel by CI** - see below.
- The abap2UI5 frontend is built to own its page; embedded, it still shows
  its busy indicator and may set the document title (embed control README,
  "Known limitations").

## What is tested

| | How |
|---|---|
| Host page boots, Office.js first, history methods restored, strict CSP | Playwright, Chromium, `test/e2e/` - against `test/e2e/office-stub.js` |
| The embedded app renders (UI5 1.136 and the 1.71 floor) | the same, with the BSP pages exactly as abapGit delivers them |
| Export writes a table into a new sheet (values, number formats, leading zeros, dates) | the stub records every `Excel.run` call |
| Take selection, Follow selection, an Excel error, outside Office, a bad `app` | the same |
| The manifest: well-formed, Microsoft's reader, element order, every `resid`, option checks | `npm test` (offline); Microsoft's validation service in CI (online) |
| The BSP: abapGit format, page directory, ICF file names, generated = committed | `npm test`, `npm run bsp:check` |
| The demo app | abaplint (Standard, Cloud, 7.02), the abap2UI5 linter |

**Not verified - needs a real Excel and a Microsoft 365 tenant:** that Excel
on Windows, Mac and the web load the add-in from the manifest; the real
Office.js against the host page (its CSP needs, the history workaround);
the real Excel API behaviour (`worksheets.add`, tables, number formats,
selection events) beyond what the stub models; logon in the task pane;
`frame-ancestors` and cookies in Excel on the web; the ICM rewrite example;
the approuter example; the icons through Office's cache; central deployment.

The smoke test runs on `@abap2ui5/node-runtime` 1.146.0, which predates
`z2ui5.cc.ExcelBridge`: the test server then answers the control's module
with a copy (`test/e2e/vendor/ExcelBridge.js`, from the commit of #2847).
With a runtime that has the control, the copy is never asked for -
`E2E_BRIDGE=bundle npm run e2e` makes sure of it (see
[Development](#development)).

## Development

```bash
npm ci
npm run check        # ESLint, Prettier, Node tests, bsp:check, abaplint (Standard, Cloud), abap2UI5 linter
npm run e2e:build    # the demo app transpiled onto @abap2ui5/node-runtime
npm run e2e          # Playwright smoke test, UI5 1.136 and 1.71
npm run e2e:serve    # the test server, to look at: http://127.0.0.1:4310/
```

- `webapp/` is the host page; `src/02` is **generated** from it and from the
  embed control in `node_modules` - edit `webapp/`, run `npm run bsp`, commit
  both. `npm run bsp:check` fails on a stale `src/02`.
- `npm run downport && npm run abaplint:702` checks the 7.02 downport - it
  rewrites `src/` in place, so run it in CI or on a throwaway copy.
- Against an abap2UI5 checkout instead of the npm runtime: in abap2UI5,
  `npm run pack:node-runtime`; here,
  `npm i --no-save <abap2UI5>/npm-package/abap2ui5-node-runtime-<v>.tgz @abaplint/transpiler-cli@<the version its package.json names>`,
  then `npm run e2e:build && E2E_BRIDGE=bundle npm run e2e` (and `npm ci`
  afterwards).

AGENTS.md has the rules for changes; `changelog.txt` the history.

## License

MIT - see [LICENSE](LICENSE).
