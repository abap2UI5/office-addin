// The smoke test of the add-in's host page: the BSP of src/02 as delivered,
// Office.js replaced by test/e2e/office-stub.js, the demo app of src/01 on
// abap2UI5 in Node. It proves that the page boots, the embedded abap2UI5 app
// renders in it, and z2ui5.cc.ExcelBridge drives the (stubbed) Excel API -
// not that a real Excel accepts the add-in: that takes Excel and an
// Microsoft 365 tenant (README, "What is tested").
import { test, expect } from "@playwright/test";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const STUB = fs.readFileSync(path.join(here, "office-stub.js"), "utf8");
const PAGE = "/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html";
const OFFICE_JS =
  /officeapis\.public\.onecdn\.static\.microsoft\/1\/office\.js/;

// host: what Office.onReady( ) reports - "Excel", or null (a browser tab)
async function open(
  page,
  { host = "Excel", query = "?app=Z2UI5_CL_XL_DEMO" } = {},
) {
  const problems = [];
  page.on("pageerror", (e) => problems.push(`pageerror: ${e.message}`));
  page.on("console", (m) => {
    const text = m.text();
    if (
      m.type() === "error" &&
      /Content.Security.Policy|Refused to/i.test(text)
    ) {
      problems.push(`csp: ${text}`);
    }
  });
  await page.addInitScript((h) => {
    window.__stubHost = h;
  }, host);
  await page.route(OFFICE_JS, (route) =>
    route.fulfill({ contentType: "text/javascript", body: STUB }),
  );
  await page.goto(PAGE + query);
  return problems;
}

// Where z2ui5.cc.ExcelBridge came from: the frontend bundle of the runtime,
// or - with a runtime older than abap2UI5 #2847 - the vendored copy the
// server answers resources/z2ui5/cc/ExcelBridge.js with (test/e2e/server.mjs).
// E2E_BRIDGE=bundle makes a run against a runtime that has the control fail
// when the copy was used.
function watchBridge(page) {
  const seen = { vendored: false };
  page.on("request", (r) => {
    if (r.url().endsWith("/resources/z2ui5/cc/ExcelBridge.js"))
      seen.vendored = true;
  });
  return seen;
}

const strip = (page, text) =>
  page.locator(".sapMMsgStrip").filter({ hasText: text });

test("boots in Excel and renders the embedded app", async ({ page }) => {
  const bridge = watchBridge(page);
  const problems = await open(page);
  await expect(page.getByText("000000000000001001")).toBeVisible();
  // ExcelBridge is loaded with the view; available turns true once it found Excel
  await expect(page.locator('[id$="--followSwitch"]')).not.toHaveClass(
    /sapMSwtDisabled/,
  );
  test.info().annotations.push({
    type: "ExcelBridge",
    description: bridge.vendored
      ? "vendored copy (test/e2e/vendor)"
      : "frontend bundle",
  });
  if (process.env.E2E_BRIDGE === "bundle") expect(bridge.vendored).toBe(false);
  await expect(page.getByText("Gearbox GX-400")).toBeVisible();
  await expect(page.locator("#z2ui5xl-message")).toBeHidden();
  // office_init.js put back what Office.js took away
  expect(await page.evaluate(() => typeof window.history.pushState)).toBe(
    "function",
  );
  expect(await page.evaluate(() => typeof window.history.replaceState)).toBe(
    "function",
  );
  expect(problems).toEqual([]);
});

test("Export to Excel writes the materials as a table into a new sheet", async ({
  page,
}) => {
  const problems = await open(page);
  await expect(page.getByText("000000000000001001")).toBeVisible();
  await page.getByRole("button", { name: "Export to Excel" }).click();
  await expect(strip(page, "Written to 'Materials 1'!A1:G11")).toBeVisible();

  const excel = await page.evaluate(() => window.__excel);
  expect(excel.sheets).toEqual(["Sheet1", "Materials 1"]);
  expect(excel.tables).toEqual([{ sheet: "Materials 1", hasHeaders: true }]);
  const [write] = excel.writes;
  expect(write.values[0]).toEqual([
    "Material",
    "Description",
    "Type",
    "Quantity",
    "Unit",
    "Price (EUR)",
    "Created",
  ]);
  expect(write.values).toHaveLength(11);
  // text stays text (leading zeros), numbers stay numbers, a date is a date
  expect(write.values[1][0]).toBe("000000000000001001");
  expect(write.numberFormat[1][0]).toBe("@");
  expect(write.values[1][3]).toBe(1250);
  expect(write.numberFormat[1][3]).toBe("#,##0.000");
  expect(typeof write.values[1][6]).toBe("number");
  expect(write.numberFormat[1][6]).toBe("yyyy-mm-dd");

  // the next export goes into a sheet of its own
  await page.getByRole("button", { name: "Export to Excel" }).click();
  await expect(strip(page, "Written to 'Materials 2'!A1:G11")).toBeVisible();
  expect(problems).toEqual([]);
});

test("Take selection reads the selected cells into the app", async ({
  page,
}) => {
  const problems = await open(page);
  await expect(page.getByText("000000000000001001")).toBeVisible();
  await page.getByRole("button", { name: "Take selection" }).click();
  await expect(strip(page, "2 rows taken from Sheet1!A1:C2")).toBeVisible();
  await expect(page.getByText("MAT-2")).toBeVisible();
  // a date cell comes back as the ISO date an ABAP d takes
  await expect(page.getByText("2024-01-02")).toBeVisible();
  expect(problems).toEqual([]);
});

test("Follow selection reports where the user selects", async ({ page }) => {
  const problems = await open(page);
  await expect(page.getByText("000000000000001001")).toBeVisible();
  const toggle = page.locator('[id$="--followSwitch"]');
  // enabled once ExcelBridge found Excel (available)
  await expect(toggle).not.toHaveClass(/sapMSwtDisabled/);
  await toggle.click();
  await expect
    .poll(() => page.evaluate(() => window.__excel.handlers.length))
    .toBe(1);
  await page.evaluate(() => window.__excel.select("Sheet1!D4:E9"));
  await expect(strip(page, "Selected in Excel: Sheet1!D4:E9")).toBeVisible();
  expect(problems).toEqual([]);
});

test("an Excel error reaches the app", async ({ page }) => {
  const problems = await open(page);
  await expect(page.getByText("000000000000001001")).toBeVisible();
  await page.evaluate(() => {
    window.__excel.failNext = {
      message: "The sheet is protected",
      code: "AccessDenied",
    };
  });
  await page.getByRole("button", { name: "Export to Excel" }).click();
  await expect(
    strip(page, "Excel: The sheet is protected (AccessDenied)"),
  ).toBeVisible();
  expect(problems).toEqual([]);
});

test("outside Office the app runs, says so, and Excel calls fail cleanly", async ({
  page,
}) => {
  const problems = await open(page, { host: null });
  await expect(page.locator("#z2ui5xl-message")).toContainText(
    "Not running in Excel",
  );
  await expect(page.getByText("000000000000001001")).toBeVisible();
  await page.getByRole("button", { name: "Export to Excel" }).click();
  await expect(
    strip(page, "Excel: Excel not available (NotAvailable)"),
  ).toBeVisible();
  expect(await page.evaluate(() => window.__excel.writes.length)).toBe(0);
  expect(problems).toEqual([]);
});

test("a bad app parameter starts nothing and says why", async ({ page }) => {
  await open(page, { query: "?app=%3Cscript%3E" });
  await expect(page.locator("#z2ui5xl-message")).toContainText(
    "is no ABAP class name",
  );
  await expect(page.locator(".sapMPage")).toHaveCount(0);
});
