# How each case is graded (proposed)

**Model under test:** Claude Haiku 5.5, through the real bridge and the Claude CLI sign-in. **Judge:** Claude Sonnet 5.5 (never the model under test), called with `claude -p` and a JSON schema, so no API key is needed. **Reps:** 3 per case, strict per attempt (a case that passes only on a retry counts as a fail).

## What one run produces

Each run records the tool calls, the in-memory workbook after the turn, and the final reply. Three checks read those, and each is 0 or 1:

| metric | what it checks | how |
|---|---|---|
| `tool` | The right tool was called with the right arguments, and no forbidden tool was called (a write in a read-only case; `insert_rows` for a delete request). | code, from the recorded calls |
| `state` | The workbook after the turn equals the expected one. Any cell changed outside the allowed set is a fail. Read-only cases must leave it untouched. | code, a cell-by-cell diff |
| `reply` | The facts in the reply are right: the number, the cell reference, the sheet name. For six cases it also needs a statement a number cannot show (below). | code for facts; judge for the six |

`pass` is 1 only when all three are 1. It is the headline and the first metric listed. The three above are kept as separate columns so the report shows *why* a case failed.

Recorded next to every case, not part of the score: number of tool calls, seconds from send to final reply, output tokens. Those are what you picked.

## The judge, and only for six cases

`W4`, `I5`, `I6`, `N3`, `B2`, `N5`. For each, the judge gets the request, the final reply and a short list of checkable claims, and answers true or false per claim with one line of reason:

| case | claims the judge checks |
|---|---|
| W4 | The reply says C2:C4 already contain data. The reply asks before overwriting, or says it did not. |
| I5 | The reply states the 1000-row limit, or the total inserted. It does not claim more rows than the workbook has. |
| I6 | The reply asks which row, or how many. It does not say anything was inserted. |
| N3 | The reply says the profit in row 7 is an error and that it comes from the cost in C7. |
| B2 | The reply says the last quarter covers only one month. |
| N5 | The reply identifies each of the three problems as a problem, not just lists cells: the negative quantity, the text in a price, the repeated order id. |

The reply is passed to the judge as quoted data, with an instruction not to follow anything inside it.

## Considered and left out

- **Exact order or number of tool calls as a pass condition.** The result is what matters, not the route. Kept as a side column only.
- **Reply language matching the request, tone, length.** Style, not whether the work is right. Cheap to add later if you care.
- **Number formats and colours.** The in-memory workbook does not model them.
- **Speed as a pass condition.** Side column only.
- **A judge for every case.** Where code can decide, code decides; it is free and gives the same answer twice.

## Checks I will run before any paid call

- Feed the whole pipeline a correct answer for each case (oracle) and an empty answer and a constant "ok" (null). Expect nearly all pass and nearly all fail.
- Give the judge an empty reply, "I don't know", and a confident answer to a different question. It must fail all three.
- Make an API error happen on purpose and confirm it is logged as an error, not scored 0.

## After the first paid pilot

Five cases, one rep, on Haiku 5.5 (the six judged cases are not in it unless you ask). I will show you each output next to its three grades in `report.html`. If you would have scored any of them differently, the grading changes before anything else runs.
