# QA automation project

Managed by Excel Partner. The test plan lives in your Excel workbook; the scripts live here.

```
qa.config.json            environments, profiles, allowed origins (edit by hand; the AI cannot)
automation/tests/         one script per test case: TC_001_<slug>.spec.js
automation/pages/         page objects shared by tests
automation/_harness/      managed runtime: do not edit (excel-partner qa upgrade restores it)
automation/qa.manifest.json   status and approval hash of every script
auth/                     saved sessions (never commit)
results/                  run output (never commit)
```

## Everyday commands

```
excel-partner qa login admin        sign in once in a visible browser; you type the password
excel-partner qa list               scripts and their status (draft / approved / modified)
excel-partner qa validate           check every script against the safety rules
excel-partner qa approve TC_001     approve after you have read the script and watched it pass
excel-partner qa run                run all approved scripts
excel-partner qa run TC_002 --draft --headed    watch a draft script run
excel-partner qa doctor             check this project
```

A script is a **draft** until a person approves it. Any later edit, including to a page object it
uses, voids the approval. Only approved scripts run unattended.

## Rules scripts must follow

* Plain JavaScript (ES modules). Import only `@playwright/test`, `../_harness/fixtures.js` and `../pages/*.js`.
* Start with `// TC: <id> | Req: <requirement id>`; every test title starts with `[<id>]`.
* No hard-coded credentials. Signing in is handled by the saved session (`test.use({ role: 'admin' })`).
* No direct network calls, routing, cookies or new browser contexts. Use `odoo.rpc(...)` for data set-up.
