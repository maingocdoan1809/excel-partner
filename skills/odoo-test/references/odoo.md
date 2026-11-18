# Odoo, as a test target

Odoo is not a generic web app. Read this when a step surprises you.

## After login
- The landing page differs by version: `/web` (16-17), `/odoo` (18). `oqa connect` records the real one in knowledge (`kind:login`); do not assume.
- A database selector (`/web/database/selector`) comes first on multi-database servers: the workspace needs its `db`.
- Two-factor (`input[name=totp_token]`) stops a script. Ask the user (`oqa connect --headed`) or use a test account without it.
- The client is a single page: `o.ready()` (done by every helper) waits for `.o_web_client` and the loading indicator; never sleep.

## Where things are
| You want | Look for | Helper |
|---|---|---|
| an app / menu | apps toggle `.o_navbar_apps_menu`, sections `.o_menu_sections`, items in `.dropdown-menu` | `o.menu("Sales", "Orders")` |
| a button | `button` by its text, inside `.o_control_panel`, `.o_form_view`, or a `.modal` | `o.button("Confirm")` |
| a form field | `.o_field_widget[name='<technical name>']` (use the technical name, labels change with language) | `o.field(name)`, `o.fill`, `o.m2o` |
| a many2one | type, then click the suggestion (`.o-autocomplete--dropdown-item` in 17+) | `o.m2o("partner_id", "Azure")` |
| list rows / pager | `.o_data_row`, `.o_pager_value` / `.o_pager_limit` (80 rows per page by default) | `o.rows()`, `o.pager()` → `(first, last, total)` |
| a message | `.o_notification` (error: `.border-danger`), dialogs `.modal` | `o.toasts()`, `oqa` diagnoses on failure |
| invalid / required fields | `.o_field_invalid`, `.o_required_modifier` | in the failure's `page says` |

When a selector does not fit this Odoo: `oqa probe --filter <word>`, use what it prints, then `oqa knowledge learn --kind locator`.

## Things that fail "for no reason"
| Symptom | Usually | Do |
|---|---|---|
| Save does nothing | a required field is empty (`.o_required_modifier`), or a validation toast was missed | read `page says`; fill it; `o.save()` raises with the page's message |
| Button not found | it lives in the Action/Print dropdown, a statusbar, or appears only in a state (draft / confirmed) | `oqa probe --filter <text>`; check the record's state first |
| "Discard changes?" dialog | you navigated away from a dirty form | save or discard first |
| Click lands on nothing | an onchange re-rendered the form | act again after `o.ready()` |
| The list is empty | the default filter hides the records, or the data is not there | clear the filter (`o.search("")`), then `o.reach(...)` to create data |
| Pagination case on a small database | fewer than 80 records | `o.reach(lambda: o.pager()[2] > 80, make)`; every record named with `o.name(...)` |
| Menu missing for a profile | permissions: that is often the thing being tested | assert it absent, with a framed picture of the menu |
| Labels differ from the workbook | the user language is not the workbook's | match by technical name; note it in knowledge |
| Read-only after save | the record's state locks fields | check the state before editing |

## Data rules
- Existing records are read-only. Create only records whose name starts with the workspace's test prefix (`o.name(...)`), through the UI.
- Creating data is a write: if the case does not need to create it, do not.
- Anything that sends mail, posts payments or calls an outside service is out of scope unless the user said so: ask.
