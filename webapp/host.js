// The host page of the abap2UI5 Excel add-in - the task pane's document.
//
// Office.js has to run in the top document of the task pane, and the
// system's own abap2UI5 page loads no foreign script (and locks itself with
// frameOptions trusted). So the add-in brings this page: Office.js in its
// head, the system's UI5, and the abap2UI5 app embedded with
// z2ui5.embed.Container (@abap2ui5/embed-control), which loads the
// frontend from <endpoint>?z2ui5-bundle of this same server. The frontend
// then runs in this window - and z2ui5.cc.ExcelBridge, the frontend's
// control for the workbook, finds window.Office and window.Excel here.
//
// URL parameters (the add-in manifest's SourceLocation carries them):
//   app        the ABAP class to start (default Z2UI5_CL_XL_DEMO)
//   endpoint   the abap2UI5 service on this server (default /sap/bc/z2ui5);
//              the control refuses anything that is not a path here
//   any other  handed to the app as a startup parameter
//              (client->get( )-t_comp_params), except the sap-* parameters
//              of the system and the _host_Info Office appends
// The app also gets office_host and office_platform - what Office.onReady( )
// reported, empty outside Office.
//
// Outside Office (a plain browser tab, for testing) the same page runs the
// same app, with a message that the workbook is not there; ExcelBridge then
// answers write( ) and read( ) with OnError. UI5 1.71 is the floor: nothing
// here is newer.
sap.ui.define(["z2ui5/embed/Container"], (Container) => {
  "use strict";

  const DEFAULT_APP = "Z2UI5_CL_XL_DEMO";
  const DEFAULT_ENDPOINT = "/sap/bc/z2ui5";
  // Office.onReady( ) never settles when Office.js is blocked half-way
  // (Microsoft, "Initialize your Office Add-in") - the app starts anyway
  const OFFICE_TIMEOUT_MS = 8000;
  // an ABAP class name, with a namespace like /ABC/ - 30 characters at most
  const CLASS_NAME = /^(\/[A-Z0-9_]{1,10}\/)?[A-Z0-9_]{1,30}$/;
  // not handed to the app: the page's own, the system's and Office's
  const RESERVED = /^(app|endpoint|sap-.*|_host_info)$/i;

  function byId(id) {
    return document.getElementById(id);
  }

  // One line above the app - plain DOM and textContent, so it shows even
  // when UI5 or the frontend could not start, and never as markup.
  function showMessage(text, kind) {
    const box = byId("z2ui5xl-message");
    if (!box) return;
    byId("z2ui5xl-message-text").textContent = text;
    box.className = `z2ui5xlMessage z2ui5xlMessage-${kind}`;
    box.hidden = false;
  }

  function wireMessageClose() {
    const close = byId("z2ui5xl-message-close");
    if (!close) return;
    close.addEventListener("click", () => {
      byId("z2ui5xl-message").hidden = true;
    });
  }

  // What Office.onReady( ) reports: { host, platform } - both null outside
  // Office, and null altogether when Office.js is not on the page.
  function officeInfo() {
    const ready = window.z2ui5xlOfficeReady || Promise.resolve(null);
    const timeout = new Promise((resolve) => {
      setTimeout(
        () => resolve({ host: null, platform: null, timeout: true }),
        OFFICE_TIMEOUT_MS,
      );
    });
    return Promise.race([ready.then(null, () => null), timeout]);
  }

  function readParams() {
    const query = new URLSearchParams(window.location.search);
    const params = {};
    query.forEach((value, name) => {
      if (RESERVED.test(name)) return;
      if (params[name] === undefined) params[name] = value;
      else params[name] = [].concat(params[name], value);
    });
    return {
      app: (query.get("app") || DEFAULT_APP).trim().toUpperCase(),
      endpoint: (query.get("endpoint") || DEFAULT_ENDPOINT).trim(),
      params,
    };
  }

  function excelHost(info) {
    const Office = window.Office;
    const excel = Office && Office.HostType && Office.HostType.Excel;
    return Boolean(info && info.host && excel && info.host === excel);
  }

  function start() {
    wireMessageClose();
    const page = readParams();
    if (!CLASS_NAME.test(page.app)) {
      showMessage(
        `'${page.app}' is no ABAP class name - check the URL ` +
          "parameter app in the add-in manifest.",
        "error",
      );
      return;
    }
    officeInfo().then((info) => {
      if (!window.Office) {
        showMessage(
          "Office.js could not be loaded - the app runs, but " +
            "without the workbook. Is officeapis.public.onecdn.static.microsoft " +
            "reachable and allowed by the Content-Security-Policy?",
          "warning",
        );
      } else if (info && info.timeout) {
        showMessage(
          "Office did not answer - the app runs, but the workbook " +
            "may not be reachable.",
          "warning",
        );
      } else if (!excelHost(info)) {
        showMessage(
          "Not running in Excel - the app runs, but without the " +
            "workbook. Open it from the add-in's task pane in Excel.",
          "info",
        );
      }
      const params = Object.assign({}, page.params, {
        office_host: (info && info.host) || "",
        office_platform: (info && info.platform) || "",
      });
      const container = new Container("z2ui5xlContainer", {
        app: page.app,
        endpoint: page.endpoint,
        params,
        width: "100%",
        height: "100%",
      });
      container.attachComponentFailed((event) => {
        const reason = event.getParameter("reason");
        showMessage(
          `The app ${page.app} could not start: ` +
            `${(reason && reason.message) || reason}`,
          "error",
        );
      });
      container.placeAt("z2ui5xl-app");
    });
  }

  start();
});
