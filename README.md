# interior-app

Apps Script based interior consulting tools collected in one repository.

## Structure

- `quote_app`: admin quote authoring and template/material management
- `prequote_app`: public survey intake and prequote recommendation flow
- `shared-docs`: shared domain rules, enums, and workflow documents
- `codex-skills`: local Codex helper skill definitions used in this workspace

## App layout

Each Apps Script app keeps deployable files in its own directory.

- `Code.*.js` / `Code.js`: server-side entry points and domain logic
- `BaseLib.gs`: shared integration helpers for cross-app access
- `*.html`: Apps Script HTMLService templates
- `appsscript.json`: manifest for clasp-driven sync where present

## Notes

- `prequote_app` server logic is split by domain section for safer maintenance.
- `quote_app` keeps the core runtime in `Code.js` and isolates late-added admin utilities in `Code.extras.js`.
- `Code.legacy.bak` in `prequote_app` is a non-deploy backup of the original monolithic file.
