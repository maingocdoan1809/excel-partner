---
name: odoo-test
description: Test an Odoo system from a testcase workbook (.xlsx). Reads the cases, judges whether they are clear and complete, plans, gets the user's approval, runs every case in a browser with a red-framed screenshot of every step, repairs failures, writes an Excel evidence file per case. Use when the user asks to test Odoo, run testcases from a file, or invokes /odoo-test. Needs Python 3.9+.
hooks:
  PreToolUse:
    - matcher: "Bash|Write|Edit|MultiEdit|Read|Grep"
      hooks:
        - type: command
          command: "python \"${CLAUDE_PLUGIN_ROOT}/hooks/guard.py\""
  Stop:
    - hooks:
        - type: command
          command: "python \"${CLAUDE_PLUGIN_ROOT}/hooks/stop.py\""
---

# odoo-test

Everything goes through one command: `python "${CLAUDE_SKILL_DIR}/oqa.py" <command>` (written `oqa <command>` below). `oqa.py` is the file next to this SKILL.md: if the path before it was not filled in with a real folder, use the folder this file is in. Every command checks the session's phase and
refuses out of turn, naming what to run instead: **follow its `Next:` line, do not improvise the order.** Reply to the user in the workspace's language (`oqa ws show`).
Files and formats: [references/contract.md](references/contract.md). Odoo specifics: [references/odoo.md](references/odoo.md). Repair: [references/repair.md](references/repair.md).

## The process

| Phase | You do | Command |
|---|---|---|
| setup | check the machine, then the workspace | `oqa doctor` (`--fix` if not READY) · `oqa init --status`; if not configured run `oqa init` **in the background**: it serves the form and opens it in the user's browser by itself (do not ask them to open anything), then exits when they save |
| INTAKE | read the workbook the user named (add-in: the open file; terminal: the file they gave) | `oqa session new --file <path>`, then at once `oqa dashboard` **in the background**: it serves the progress page and opens it on this session (or reopens the running one). Do it without asking |
| ANALYZE | judge every case, then hand it in | `oqa analyze --template --out <scratch>/analysis.json` → fill → `oqa analyze --submit <file>` |
| CONNECT | sign every profile in | `oqa connect` (`--headed` if a two-factor code is needed: ask the user) |
| CONFIRM | build the strategy, show it, **stop** | `oqa strategy`. Tell the user the strategy is waiting on the dashboard that is already open. Only they confirm (the Confirm button there, or `oqa confirm` in their terminal). You cannot, and must not try. Then `oqa wait` **in the background** |
| EXECUTE | explore each step, build the script, run | `oqa scaffold` → `oqa explore` / `oqa try` → `oqa build` → `oqa run` |
| REPORT | close | `oqa report`: it writes ONE Excel evidence workbook (a summary sheet, then a sheet per case: what the testcase asked, how you understood it, every step with its red-framed picture, the runs) and **opens it on the user's screen**. Tell them the file path. They can also rebuild or download it any time from the dashboard |

Anytime: `oqa status` (also lists cases whose failure is not explained yet) · `oqa repair ID …` · `oqa ask "<question>" [--case ID]` (then `oqa wait` in the background and stop) · `oqa wait` (ends when the user acts on the dashboard: confirms, answers, sends a hint, stops or resumes, and says which; the dashboard has a Stop button, and a hint box for a case that keeps failing) · `oqa knowledge search|learn` · `oqa probe [--filter text | --failed]` · `oqa lookup models|fields|xmlid|find|search …` · `oqa runner status|stop` · `oqa dashboard`.

## ANALYZE: the part that decides everything

The workbook is written for people: short, with assumed data. For **each case** fill the template honestly:
- `clarity` 1-5 (is the intent unambiguous?) and `sufficiency` 1-5 (does it carry what is needed to run?). Below 3 means `needs_user`: put the questions in `questions` and `oqa ask` them now, not after a failed run.
- `refs`: every file or record the case points at. Intake already marked which files exist; a missing one is `needs_user`, never invented.
- `data_needs`: data the environment may lack (a list long enough to paginate, a second company, a record in a given state), each with a `fallback`: the step that creates it (`o.reach`, `o.new` + `o.save`). A case must run on **any** environment, so plan the creation, do not hope.
- `edge_cases`: what can differ between environments or runs: empty list, more than one page, permissions of the profile, language, default values, a record already existing.
- `depends_on` + `reuse`: when a case can start from where another ended, say so; the runner then keeps the screen instead of starting over.
- `steps`: what you will actually do, finer than the workbook, each with an observable `expected`. They become the captions of the evidence pictures.
- `previous` (in the template, only when this file was tested before): how the case ended last time, why, and the user's hints. **Use it**: follow the hints, plan for the reason it failed (a precondition to create, a selector that did not fit), do not repeat an approach that failed, and do not ask again what the user already answered.
- `profile`: which profile (Admin, Teamlead, …) runs it.
- `timeout_s` (optional, 5-120, default 8): for a case whose pages are slow (a report, a wizard, a long list).
- Read each case the way the product works, not only as written: what must be true for a step to work (a required field, a state, a permission, a record that must exist)? Put it in `data_needs` / `refs`, or `needs_user`.
- A phrase with two readings that give different results ("the customer who bought the most": one big order, or the largest total across orders?): write the reading you chose into the step, and ask when the outcomes differ.

## EXECUTE: one case at a time; each step is done the moment it is verified

Work **case by case**, never "explore every case, then build, then run". For ONE case: explore its steps, `oqa build`, `oqa run ID`, repair if needed, then move to the next case. Do not explore a case you are not about to finish. A step whose `oqa try` says **ok** has been done for real on the page: go straight on to the next step, do not stop to review other steps or cases first.

A selector is never a guess. Every line of a script ran once on THIS Odoo, before it was written down, and a wrong guess found while exploring costs nothing. `oqa run` refuses a script that was not built this way.

1. `oqa scaffold` lists the cases (status `reuse` = scripted before and still valid: **do not open it**; `new` / `changed` = explore it). Take them in order, one at a time.
2. **Where is it?** Before exploring, if the workspace has a sitemap (`oqa sitemap show`), ask it instead of clicking around: `oqa sitemap find <words> [--kind menu|action|view]` (a menu, a model, a button label), `oqa sitemap buttons <screen> [--match <word>] [--view form]` (the buttons with their `[name=...]`, a selector to start from and where they lead), `oqa sitemap fields <screen> [--required] [--match <word>] [--view list]` (the form's fields with type, required / read only, selection values and a selector; ask for the few you need, it prints at most 40 lines), `oqa sitemap path <target> [--from <screen>]` (the way there); the MCP tools `sitemap_find`, `sitemap_buttons`, `sitemap_fields`, `sitemap_path` answer the same. `oqa sitemap build` makes it (read only, clicks nothing; once per Odoo, again after a module is installed; if the user has none, offer it). Never open `knowledge/sitemap.json` itself (it is large): ask with these, narrowest first. It tells you where to go and what the form holds; its selectors are where to start, and only `oqa try` proves a step works. A `type="object"` button leads nowhere on the map: that is `unresolved` by nature.
   Sitemap queries search both English and Japanese labels. `oqa sitemap build --languages en,ja` requires both languages to be active in Odoo; without it, the builder reads whichever are active and reports missing languages. Use `--lang en|ja` on sitemap commands (MCP argument `lang`) for labels and explanations in that language; by default the workspace's English/Japanese language is used, with English for Vietnamese workspaces. Rebuild an old sitemap to collect translations. Technical names and selectors remain unchanged. See [the sitemap contract and examples](references/contract.md#sitemapjson).
3. **Explore**: `oqa explore ID` opens a new page where step 1 begins and prints what is on it (the controls, with the technical names `[name=...]`). For each step, in order: `oqa try ID --step N --code "<python>"` (or `--file`). The code runs for real, with `t` `o` `s` `page` `t.data` in scope, the same as in the script. It works or it does not:
   - The page is read only after it has **settled** (no spinner, and the controls and text unchanged for ~0.4 s), so records and widgets that render late are already there. If the printed page still looks half-loaded or the output says it did not settle, `oqa probe --filter <word>` once more before writing a selector; never write a selector from a loading screen, and never `sleep`.
   - **ok**: it is kept as a verified piece of the step, and the page it left is printed so you can write the next piece from what is there (`--filter <word>` lists only matching controls).
   - **FAILED**: nothing is kept, no attempt is used, and you see the page as it is plus the controls closest to what the code asked for. Change the code and try again.
   Write each step as pieces that end with `s.hl("<selector>")` (the element that gets the red frame; it must exist when tried) and a check of the expected result (`s.check`, `s.equals`, `o.expect_*` or `assert`): content the step names (a value, a count, a status), not only that something is on screen. Steps pass values to each other through `t.data`, never through variables. Record names come from `o.name("...")` (or `o.unique(...)` when the model refuses a duplicate name). Never weaken an `expected`. `--reset` starts a step over, `--undo` drops its last piece, `oqa explore ID --to N` takes a new page to step N by replaying what is verified (no tokens spent).
4. `oqa build ID` writes `scripts/<workbook>/<case>.py` from the verified pieces and says which steps are not tried yet or have no check / `s.hl`. Do not edit the script by hand: an edited script is refused.
5. `oqa run ID` runs that case (plain `oqa run` runs the remaining ones) in one browser and **stops at the first failure**, printing the failed step, what the page said, the picture, the controls closest to what it wanted and known fixes. The runner stays alive between runs: after a repair, `oqa run ID` **picks up at the step that failed**, with the steps before it kept, on the page where it stopped. A run that gets further than the one before does not use up a repair; the same step failing again does. You get the first run plus **3 repairs** per step; the runner refuses the next one.
6. **A failed run is explained before it is rerun.** The page where the step stopped is saved: `oqa probe --failed --case ID --filter <word>` reads it with no browser. Then `oqa explore ID --to N` and `oqa try ID --step N --reset ...` until it works, `oqa build ID`, and record what you found: `oqa repair ID --cause "…" --evidence "…" --change "…"` (`--reanalysis "…"` when the same failure comes twice; `--timeout N` when the page is only slow). `oqa run` refuses a rerun without that note, and `repair` refuses a cause you already gave. The user sees each note on the dashboard as how you re-read the case. A run after exploring starts from step 1 (the page moved) and proves the whole case again.
7. A `user hint:` line in `oqa run` output is the user's advice for that case: read it first and follow it. After a hint the case is back in the queue with a fresh budget.
8. When the budget is spent or the cause is not yours to fix: `oqa ask --case ID "…"` (needs input), or `oqa finish ID --status FAIL|BLOCKED|PENDING --reason "…"`. FAIL is allowed only after a run that failed with a picture. PASS only comes from a run.
9. After every repair that taught you something about this Odoo: `oqa knowledge learn --screen <model or menu> --kind <kind> --fact "…"`.

## Spend tokens like they cost money

- Read what commands print; do not open session files, whole pages or scripts marked `reuse`.
- Look at a picture only when the printed reason and `page says` are not enough.
- A model, field or external id is never guessed from a label on screen: `oqa lookup models <text>`, `oqa lookup fields <model> --filter <label>`, `oqa lookup xmlid <module.name>`, `oqa lookup find <model> --filter <name>` (read only, the profile's own rights; the same as `o.models`, `o.fields`, `o.xmlid`, `o.find`, `o.records`, `o.model_exists` in a `try`).
- Before debugging: `oqa knowledge search --text <words>`. To find a selector: `oqa probe --filter <word>` (a few lines), not a page dump.
- Same failure twice with no new information: change the approach, do not rerun it. Exploring first is cheaper than a failed run.

## Never

Edit files under the oqa Home by hand · read signed-in state or passwords · confirm for the user · run a case script outside `oqa run` · call a case PASS yourself · test data without the workspace prefix.
