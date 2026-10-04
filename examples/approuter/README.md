# The host page behind an SAP approuter

An example - not tested by CI - for a system that has no BSP applications
(SAP BTP, ABAP environment) or a landscape that reaches the system through
an approuter anyway. The approuter serves the host page and forwards the
abap2UI5 service, so both are on its origin, which is what the add-in
needs (README, "Same origin").

- `dist/` next to `xs-app.json`: `npm run dist` in the repository root
  writes it (the host page and the embed control).
- Destination `abap2ui5`: the ABAP system. On SAP BTP, ABAP environment
  the abap2UI5 service is a custom HTTP service with a path of its own -
  route that path, and write the manifest with
  `--endpoint <that path> --path /index.html --host https://<approuter>`.
- Destination `ui5`: `https://ui5.sap.com` (or `https://sdk.openui5.org`),
  for `resources/` of the host page - the version in the target is the UI5
  your apps run on. Allow that origin in the Content-Security-Policy instead
  of the system's own UI5.
- Excel on the web frames the approuter's page: the approuter must not send
  `X-Frame-Options`, and the Content-Security-Policy it sends (for example
  through the approuter's `httpHeaders` environment variable) needs the
  `frame-ancestors` of the main README.
- abap2UI5 fetches a CSRF token when the approuter's route asks for one
  (abap2UI5 1.146.0 and later).
