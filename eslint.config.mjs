import js from "@eslint/js";
import globals from "globals";

const rules = {
  ...js.configs.recommended.rules,
  "no-unused-vars": [
    "error",
    { caughtErrors: "none", argsIgnorePattern: "^_" },
  ],
  eqeqeq: ["error", "smart"],
  "prefer-const": "error",
};

export default [
  {
    ignores: [
      "**/node_modules/**",
      // the generated BSP (scripts/build-bsp.mjs) and the vendored control
      "src/**",
      "dist/**",
      "test/e2e/vendor/**",
      "test/e2e/.build/**",
      "test/e2e/.deps/**",
      "test-results/**",
      "playwright-report/**",
    ],
  },
  // the host page: browser scripts and one UI5 module, on the UI5 1.71 floor
  {
    files: ["webapp/**/*.js", "test/e2e/office-stub.js"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: { ...globals.browser, sap: "readonly" },
    },
    rules,
  },
  // tooling and tests
  {
    files: ["**/*.mjs"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: { ...globals.node, ...globals.browser },
    },
    rules,
  },
];
