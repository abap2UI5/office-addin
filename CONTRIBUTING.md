# Contributing

Thank you for helping. Issues and pull requests are welcome at
https://github.com/abap2UI5/office-addin.

- Read [AGENTS.md](AGENTS.md) first: it says what belongs here and what
  belongs to abap2UI5 (the `z2ui5.cc.ExcelBridge` control, the frontend) or
  to abap2UI5/embed-control.
- `webapp/` is the source of the host page; `src/02` is generated from it -
  run `npm run bsp` and commit both.
- Before a pull request: `npm ci && npm run check`, and for a change to the
  host page or the demo app also `npm run e2e:build && npm run e2e`. CI runs
  these, the 7.02 lint and Microsoft's manifest validation.
- Something you verified in a real Excel or Microsoft 365 tenant is worth a
  note in the pull request - CI cannot do that, and the README's "What is
  tested" should say so.
- English, an imperative subject line, one topic per pull request.
