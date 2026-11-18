# Contract: files and commands

This is everything another program (the Office Partner add-in, a dashboard, a script) may rely on. Nothing else in `oqa/` is a public interface.
Every JSON file has `schema_version` (now 1). A reader that meets a larger number must stop and tell the user to update the skill.
`oqa --version` prints the CLI version. Every command accepts `--json`. Exit codes: 0 done, 1 unexpected failure, 2 refused (wrong phase, invalid input, waiting for the user).

## Home

`$ODOO_TEST_HOME`, else the folder in `~/.odoo-test.json` (`{"home": "..."}`), else `~/.odoo-test`.

```
<home>/workspaces/<ws>/workspace.json
<home>/workspaces/<ws>/knowledge/index.md
<home>/workspaces/<ws>/knowledge/sitemap.json             the screens of this Odoo and how they lead to each other (`oqa sitemap build`)
<home>/workspaces/<ws>/current                       id of the session used when none is named
<home>/workspaces/<ws>/sessions/<id>/
    state.json  cases.json  analysis.json  strategy.json  strategy.md  events.jsonl  evidence/  logs/
```

A session is chosen by `--ws` / `--session`, else `$OQA_WS` / `$OQA_SESSION` (what the add-in sets for a chat), else the workspace's `current`.

## workspace.json

`{schema_version, name, url, db, profiles:[{name, login, note}], reference_dirs:[path], language: vi|ja|en, test_prefix}`.
Never a password: each profile's password is in the credential store under the key `oqa:<workspace>:<profile>`
(Windows Credential Manager; the `keyring` package elsewhere).

## sitemap.json

`{schema_version, meta:{generated, profile, odoo_version, menus, actions, views, object_buttons, truncated, languages?, default_language?, missing_languages?}, nodes:{key: node}, edges:[{from, to, kind, label?, name?, labels?}], unresolved:[{where, why, message?, values?, language?}]}`.
Read from Odoo's menus, actions and view definitions through read-only calls (nothing is clicked); it shows what the profile named in `meta` may see.
Node keys: `menu:<id>` · `window|server|client|url|report:<id>` (actions) · `view:<model>:<type>:<id>`. A view node carries `buttons:[{label, name, type: action|object|open|edit, invisible?, stat?}]` and `disabled:[create|edit|delete|duplicate]` (what the view switches off: no New button, read only) and `fields:[{name, label, type, required, readonly, relation?, options?, invisible?, widget?}]` (`required` / `readonly` are true, false or the expression of the view when it depends on other fields; `options` are `value=label` of a selection).
Edge kinds: `menu` (parent to child), `opens` (menu to action), `shows` (action to view), `row` (list or kanban to the form of the model), `button` (view to the action a `type="action"` button opens).
A `type="object"` button calls a method: where it leads is unknown until it is run, so it has no edge (it is counted in `meta.object_buttons`). Asked through `oqa sitemap find|buttons|path` or the MCP server `oqa mcp` (tools `sitemap_find`, `sitemap_buttons`, `sitemap_fields`, `sitemap_path`; `oqa mcp --config` prints its `.mcp.json` entry).

The builder reads active `en_US` and `ja_JP` translations using a per-call `context.lang`; it never changes the user's language or activates languages. `--languages en,ja` explicitly requires both to be active; without this option the builder uses whichever are active and reports missing languages in `meta.missing_languages`. At least one must be active.
Nodes, fields, buttons and labelled edges carry `labels:{en_US: text, ja_JP: text}` for available translations. Selection fields also carry `options_by_language:{code:[value=label]}`. The original `label` / `options` stay in the first built language (`meta.default_language`); technical keys, names, selectors and graph counts do not vary by display language.
Find, screen resolution and `--match` search every stored translation and selection label, with Unicode NFKC normalization (including fullwidth Latin and halfwidth kana) and case folding. CLI `--lang en|ja` and MCP `lang` choose both the labels and explanatory text of the response; by default use the workspace's English/Japanese setting (Vietnamese workspaces retain the existing English sitemap output). Missing label translations fall back to the original label. Older schema-1 maps remain readable; rebuild them to collect translated labels. Structured unresolved messages are translated for display; original Odoo error details are preserved.

Examples:

```sh
python oqa.py sitemap build --languages en,ja --lang ja --ws demo
python oqa.py sitemap find 注文 --lang en --ws demo
python oqa.py sitemap buttons Orders --match 確認 --lang ja --ws demo
python oqa.py sitemap fields 注文 --required --lang ja --ws demo
python oqa.py sitemap path "Cancel order" --lang ja --ws demo
```

MCP example: `sitemap_fields` with `{"screen":"Orders","match":"顧客","lang":"ja"}`. A search can use either language regardless of the output language. Labels are taken from Odoo's installed translations; untranslated Odoo terms retain Odoo's fallback text.

The progress dashboard has a workspace-level `site_map` tab. It reads the saved map (no Odoo connection or session is required) and draws an interactive SVG graph: workspace → root menus → menus → actions → views, with button and row transitions when a view is expanded. Drag to pan, scroll or use +/− to zoom, fit the graph, open/close branches, or return to the overview. A minimap shows the current viewport. Search locates a screen on the graph and highlights its known route. Menus whose parents are unavailable are kept in an explicit "Other menus" group; that display grouping never invents an Odoo navigation route. The inspector shows buttons/fields as cards and navigation as numbered steps in English/Japanese. Workspace changes clear the selected screen and filters; switching back to progress keeps the selected case. Reload refreshes the saved map after a rebuild. Workspaces without a map show an empty state.
Authenticated dashboard routes: `GET /api/workspaces` lists workspace names/languages and map availability; `GET /api/sitemap?ws=<name>&lang=en|ja&what=show|find|screen` returns summary plus compact topology (node labels, IDs, counts and edges; no field/button definitions), search results (`q`, optional `kind`, at most 60), or screen detail (`screen`, optional `view`, `match`, `required=true`). Screen details include bounded structured `controls`, `route`, `route_edges`, and `reachable`, alongside the existing CLI text. These routes use the dashboard's local token and Host/Origin checks. A direct dashboard URL can use `tab=site_map&workspace=<name>`.

## state.json

`{schema_version, id, workspace, file, created, phase, done:[{phase, at}], waiting, strategy_sha, confirmation, attempts:{case:n}, results:{case:{status, attempts}}, connect, stopped (null or {by, at}), hints:{case:[{text, at, by}]}, seen, timeouts:{case:seconds}, progress:{case:furthest failed step}, explore: "required"|"off"}`

Phases: `ANALYZE → CONNECT → CONFIRM → EXECUTE → REPORT → DONE` (INTAKE happens inside `session new`).
`waiting` is `null` or `{question, case, asked_at}`: while set, only `answer`, `status`, `ask` and reading work.
`attempts[case]` counts runs: the first plus at most 3 repairs; the next is refused. An answer about a case resets its count.

## cases.json (INTAKE)

`{schema_version, source:{file, name, sha256, sheet, header_row}, warnings:[text], cases:[{id, title, category, type, role, pre, steps:[text], expected, note, sheet, row, refs:[{kind:"file", text, status: found|missing, path}], case_refs:[id]}]}`

## analysis.json (ANALYZE)

Produced by `oqa analyze --template`, completed by the assistant, accepted only by `oqa analyze --submit`. Per case:
`id, profile, verdict (ready|needs_user|blocked), clarity{score 1-5, notes}, sufficiency{score 1-5, missing}, refs[{text, status: ok|missing|asked|waived, note}], data_needs[{what, fallback}], edge_cases[] or edge_cases_none_reason, depends_on[], reuse, steps[{n, action, expected}], timeout_s (optional, 5-120, default 8), questions[] (needs_user) / reason (blocked)`.
`ready` needs clarity ≥ 3, sufficiency ≥ 3, no reference left `missing`, and may depend only on other `ready` cases. No cycles.

## strategy.json / strategy.md (CONFIRM)

`strategy.json = {schema_version, order:[case], excluded:[{id, verdict, reason}], analysis_sha, sha}`. `strategy.md` is the same for people.
`confirm` succeeds only if `sha` is intact and matches `analysis.json`: the user approves exactly what will run.
**Only a person confirms**: `oqa confirm` needs a real terminal (the assistant's shell is not one), or the dashboard's Confirm button.

## events.jsonl

One JSON object per line, appended in order; each has `ts` and `kind`. Readers skip a line they cannot parse (it may be half written).

| kind | fields |
|---|---|
| session_started | workspace, file |
| phase | phase, text |
| ask / answered | question / answer, [case] |
| confirmed | by, strategy_sha |
| case_planned | case, title, steps[{n, action, expected}], order |
| stopped / resumed | by (the user's Stop and Resume) |
| hint | case, text (the user's advice; the case goes back in the queue with a fresh budget) |
| attempt_finished | case, attempt, status PASS\|FAIL\|STOPPED, reason, [failed_step, script_sha] |
| repair | case, attempt (the failed run), cause, evidence, change, [reanalysis, probed, script_sha, timeout_s, unchanged_reason]: the assistant's note after a failed run; the next run of the case needs one |
| probe | url, count, [case, filter]: the assistant looked at the real page |
| explore | case, step, ok, [code, url, error]: the assistant tried code for a step on the real page |
| case_started | case, attempt, [from_step]: `from_step` > 1 when the run picked up at a step that failed before (the steps before it are carried from the earlier run) |
| case_started | case, attempt |
| step | case, attempt, n, status PASS\|FAIL\|SKIPPED, caption, expected, actual, [checks], images[{raw, annotated}] (`checks` = how many checks of the expected result ran; a PASS step needs at least one) |
| case_finished | case, status PASS\|FAIL\|PENDING\|BLOCKED\|NOT EXECUTED, attempts |
| session_finished | counts |

A `step` with status PASS or FAIL **must** carry at least one image pair (`raw` and `annotated`, the red-box one). The writer refuses anything else.

## knowledge/index.md

`- [screen:<screen>] [kind:<kind>] [tag:<word>] <fact> -> <ref or ->`, one line per fact, newest last.
kind: login | navigation | locator | data | error | fix | quirk | rule. Never a password, token or cookie.

## Commands (this version)

| command | what it does |
|---|---|
| `init [--ws N] [--timeout S] [--no-open] [--status]` | serves the configuration form on 127.0.0.1 and **blocks until it is saved** (exit 0, summary without secrets) or times out (exit 2, nothing written). Run it in the background; `--status` only reports whether a workspace is configured |
| `doctor [--fix]` | checks Python, packages, browser, credential store; `--fix` installs what is missing |
| `ws show\|list` | the workspace (no secrets: each profile says whether a password is saved) |
| `session new --file F` / `show` / `list` | INTAKE: read the workbook → `cases.json`, session starts in ANALYZE |
| `analyze --template [--out F]` / `--submit F` | the skeleton to fill / hand in the analysis (validated; ANALYZE → CONNECT) |
| `strategy` | order the ready cases, write `strategy.json` and `strategy.md` (phase CONFIRM) |
| `confirm` | the user approves (terminal only); CONFIRM → EXECUTE |
| `ask "Q" [--case ID]` / `answer "text"` | wait for the user / record their reply |
| `wait [--timeout S]` | blocks until the user acts on the dashboard (confirm, answer, hint, stop, resume), then says what; run it in the background |
| `stop` / `resume` | the USER stops the test (the step in progress finishes) / lets it go on (`resume` needs a real terminal); also buttons in the dashboard |
| `repair ID --cause --evidence --change [--reanalysis] [--timeout S] [--unchanged why]` | after a failed run: record what was found and changed; `run` refuses a rerun without it |
| `probe [--case ID] [--filter T] [--failed]` | what can be clicked or filled on a page; `--failed` reads the page saved when a step failed (no browser); logged as a `probe` event |
| `explore ID [--to N]` / `try ID --step N --code C [--file F] [--reset] [--undo] [--filter T]` | open the page where step N begins / run code for a step on it; kept in `<script>.recipe.json` only when it worked, no attempt used. Needs the runner |
| `build ID [--force]` | write the case script from the verified pieces, sealed with `# oqa-built: <digest>`; `run` refuses a script that is not sealed (session `explore: "required"`) |
| `run [ID...] [--from N] [--fresh] [--inline]` | run in the runner (it keeps the browser between runs; `--inline` runs and closes in this process); a failed case picks up at the step that failed |
| `runner status\|stop` | the background process that keeps the browser (`runner.json` in the session folder holds its address and token) |
| `status` | phase, progress, the next command, cases whose failure is not explained yet |
| `knowledge search\|learn` | look up / add what was learned |

Coming (M2): `connect`, `run <ID>`, `report`, `dashboard`.
