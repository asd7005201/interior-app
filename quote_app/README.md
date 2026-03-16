# quote_app

Interior quote admin app for estimate editing, dashboarding, template management, and material master operations.

## Important files

- `Code.js`: main Apps Script server runtime
- `Code.extras.js`: lower-risk add-on admin helpers split from the legacy tail block
- `BaseLib.gs`: reusable cross-app access helpers
- `*.html`: admin and customer-facing HTMLService screens
- `docs/architecture.md`: high-level domain outline

## Deployment

Add a local `.clasp.json` that points to this directory and push with `clasp`.
