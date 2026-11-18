# Repair: a step failed

`oqa run` printed: the failed step, the reason, `page says` (the errors, invalid fields and dialogs the page itself shows), the picture, and any `known:` fixes. In that order:

1. **Known fix?** Apply what `known:` printed. If nothing: `oqa knowledge search --text <two words of the reason>`.
2. **Classify**, using the reason and `page says` (open the picture only if they do not settle it):

| It is | Evidence | Do |
|---|---|---|
| a selector or timing problem | "not found", "Timeout … waiting for locator", the page looks right | `oqa probe --filter <word>`, change the selector, rerun |
| missing data | empty list, "no records", fewer rows than the case needs | use the case's `data_needs` fallback: `o.reach(...)` / create through the UI with `o.name(...)` |
| a validation the case did not satisfy | `page says: fields: …` / `errors: …` | fill the named field with a valid value; the case, not the app, was incomplete |
| the application is wrong | the page works, the result differs from `expected`, and a second approach gives the same | do not "fix" it: `oqa finish <case> --status FAIL --reason "<what happened vs expected>"` (needs the failed run's picture, which you have) |
| a decision or input only the user has | a missing file, an account, a business rule | `oqa ask --case <case> "…"`, stop |
| the environment | HTTP 500 dialog, session expired, server down | `oqa finish <case> --status BLOCKED --reason "…"` |

**FAIL means the application contradicts the case under the conditions the case states.** If the step failed because the case assumed something the application requires and never said (for example: Publish needs at least one placeholder), that is a gap in the case, not a defect. Meet the condition through the UI if that is cheap and in scope (then `oqa knowledge learn --kind rule`), otherwise `oqa ask --case <case> "…"` or `oqa finish <case> --status BLOCKED --reason "testcase thiếu điều kiện: …"`. Do not leave the remaining steps unexecuted under a FAIL you did not have to give.

3. **Fix the step by exploring it**, never the case: `oqa explore <case> --to <step>` (a page where the step begins), then `oqa try <case> --step <step> --reset --code "..."` until it works (a failed try costs nothing and shows the controls closest to what you asked for), then `oqa build <case>`. The page where the failed step stopped is saved too: `oqa probe --failed --case <case> --filter <word>`. You may change selectors, waits, order of actions, data creation. You may **not** change an `expected`, delete a step, or add `try/except` to swallow a failure. Never edit the script file by hand: it is refused.
4. **Say what you found**: `oqa repair <case> --cause "<why it failed>" --evidence "<what showed it>" --change "<what you changed>"`. Without it `oqa run <case>` is refused.
   - *Cause* is a finding, not a wish ("the Save button is in the gear menu in Odoo 19", not "try another selector"). *Evidence* is what you looked at: the picture, a probe line, the page text, a knowledge entry.
   - Probe the page between the failure and the note (`oqa probe --case <case> --filter <word>`); the note records whether you did, and the user sees it.
   - A cause that repeats an earlier one is refused: the fix it led to did not work, so the cause is something else.
   - A script identical to the failed one is refused, unless you say what else changed (`--timeout N` for a slow page, `--unchanged "…"` for data made through the UI).
   - **The same failure twice** needs `--reanalysis "…"`: re-read the case and say what you assumed that was wrong (a precondition the case never stated, a different screen, a state left by another case).
5. **Rerun that case only**: `oqa run <case>`. After a run that only fixed the step that failed the runner picks up there, with the earlier steps kept; after exploring it starts from step 1 (the page moved) and proves the whole case. A run that fails at a LATER step than before is not a failed repair; the same step failing again is. You have 3 repairs per step. The runner refuses the next one.
6. A Timeout is one of two things: the element is not there (wrong selector, wrong screen: probe) or the page is slow (`--timeout 20`; never `wait_for_timeout` in a script). Decide which from the picture before choosing.
7. After a repair that worked, `oqa knowledge learn` the fact in one line (`kind:fix` or `kind:locator`).

When the budget is gone the right move is `oqa ask` or `oqa finish`, not another guess: the user would rather answer one question than wait for a fourth attempt.
