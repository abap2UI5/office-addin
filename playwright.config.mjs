import { defineConfig } from "@playwright/test";
import fs from "node:fs";

// The smoke test of the host page: test/e2e/server.mjs serves the BSP of
// src/02 as delivered, abap2UI5 (@abap2ui5/node-runtime) with the demo app,
// and UI5 from npm - one origin, like the SAP system. Office.js is the stub
// test/e2e/office-stub.js. Run `npm run e2e:build` first.
//
// Two projects, one server each: UI5 1.136 and UI5 1.71, the floor of
// abap2UI5 and of the embed control.
//
// Chromium: CHROMIUM_BIN, else /opt/pw-browsers/chromium where it exists,
// else Playwright's own (`npx playwright install chromium` in CI).
const local = ["/opt/pw-browsers/chromium"].find((p) => fs.existsSync(p));
const executablePath = process.env.CHROMIUM_BIN || local;
const PORTS = { "ui5-1.136": 4311, "ui5-1.71": 4312 };

export default defineConfig({
  testDir: "test/e2e",
  timeout: 90_000,
  expect: { timeout: 20_000 },
  fullyParallel: false,
  // one worker per server: the transpiled framework keeps per-request state
  // in static attributes, so requests of two tests must not interleave
  workers: 1,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? [["list"], ["html", { open: "never" }]] : "list",
  use: {
    // the width of a task pane in Excel on the desktop
    viewport: { width: 350, height: 800 },
    launchOptions: executablePath ? { executablePath } : {},
    trace: "retain-on-failure",
  },
  projects: Object.entries(PORTS).map(([name, port]) => ({
    name,
    use: { baseURL: `http://127.0.0.1:${port}` },
  })),
  webServer: Object.entries(PORTS).map(([name, port]) => ({
    command: `node test/e2e/server.mjs --port ${port} --ui5 ${name.slice(4)} --csp`,
    url: `http://127.0.0.1:${port}/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html`,
    reuseExistingServer: !process.env.CI,
    timeout: 90_000,
  })),
});
