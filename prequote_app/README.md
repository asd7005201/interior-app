# prequote_app

Public prequote survey app that collects customer intent and generates estimate ranges plus recommendations.

## Important files

- `Code.00.bootstrap.js` to `Code.13.triggers.js`: server logic split by domain section
- `Code.legacy.bak`: backup of the original monolithic file, excluded from deployment
- `BaseLib.gs`: quote master integration helpers
- `survey.html`, `result.html`, `admin.html`: HTMLService views
- `appsscript.json`: Apps Script manifest

## Deployment

Add a local `.clasp.json` that points to this directory and push with `clasp`.
