# -*- coding: utf-8 -*-
"""Run event stream: one JSON line per event, appended to <run out>/events.jsonl (or the file named by env QA_EVENTS).
It is the single source of progress for everyone who watches a run (Excel add-in, dashboard, a person tailing the file).
Kinds: run_planned | phase | run_started | case_started | shot | step | case_finished | run_stopped | run_finished.
  run_planned  the plan, BEFORE anything runs: every case with its steps, expected result and row in the workbook
  phase        what the agent / runner is doing now (setup, automate, preflight, run, cleanup, report ...) + a sentence
  shot         one evidence picture, the moment it is taken (a step usually has several)
Writing an event NEVER breaks a test: any error here is swallowed."""
import json, os, time
from pathlib import Path
from .config import cfg


def path():
    p = os.environ.get("QA_EVENTS")
    return Path(p) if p else Path(cfg.out) / "events.jsonl"


def emit_to(_file, _kind, /, **data):
    """Positional-only on purpose: event fields such as `file=` (the workbook of a case) must never collide with a parameter name."""
    try:
        file = Path(_file)
        file.parent.mkdir(parents=True, exist_ok=True)
        rec = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "suite": cfg.suite or ""}
        rec.update(data)
        rec["kind"] = _kind                    # the reserved name always wins over a data field of the same name
        with open(file, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass


def emit(_kind, /, **data):
    try:
        emit_to(path(), _kind, **data)
    except Exception:
        pass                                   # progress reporting must never stop a test


def describe_plan(plan):
    """Cases of a run plan with everything a person wants to know before it starts: title, steps, expected result, where it lives in the workbook."""
    from ..reporting import spec as spec_mod
    from .config import WORK
    out = []
    for item in plan.get("items", []):
        by = {}
        try:
            sp = Path(item.get("spec") or "")
            sp = sp if sp.is_absolute() else WORK / sp
            by = {str(c["no"]): c for c in spec_mod.read(sp)}
        except Exception:
            pass
        for i in item.get("ids", []):
            c = by.get(str(i), {})
            out.append(dict(suite=item["suite"], id=str(i), title=str(c.get("view", "")), pre=str(c.get("pre", "")), expected=str(c.get("exp", "")),
                            steps=list(c.get("act_lines", [])), role=str(c.get("role", "")), sheet=c.get("_sheet", ""), row=c.get("_row", 0),
                            header=c.get("_header", 0), cols=c.get("_cols", {}), file=str(item.get("spec") or "")))
    return out
