"""What the NEXT session on the same testcase file should already know: the user's hints, how each case ended last time and why. A re-test that starts from
nothing repeats the same failures and asks the same questions; this is what makes the second run better than the first.

    <home>/workspaces/<ws>/memory/<testcase file stem>.json
    {"schema_version": 1, "cases": {"<case id>": {"status", "reason", "session", "at", "hints": [{"text", "at", "session"}]}}}

A hint is remembered the moment the user sends it. A result is remembered when a case ends. A workspace that already has sessions but no memory yet is read back
from them once, so nothing the user said earlier is lost."""
import re
import time
from pathlib import Path

from . import SCHEMA_VERSION
from .paths import session_dir, session_ids, workspace_dir
from .store import locked, read_json, write_json

MAX_HINTS = 8
MAX_REPAIRS_KEPT = 6


def stem_of(file):
    return re.sub(r"[^\w.-]+", "_", Path(str(file)).stem) or "workbook"


def _file(ws, stem):
    return workspace_dir(ws) / "memory" / f"{stem}.json"


def load(ws, stem):
    data = read_json(_file(ws, stem), None)
    if data is None:
        data = {"schema_version": SCHEMA_VERSION, "cases": {}}
        backfill(ws, stem, data)
    return data


def _save(ws, stem, data):
    write_json(_file(ws, stem), data)


def _case(data, case):
    entry = data["cases"].setdefault(case, {"status": "", "reason": "", "session": "", "at": "", "hints": []})
    entry.setdefault("repairs", [])
    return entry


def remember_hint(ws, stem, case, text, session_id):
    with locked(_file(ws, stem)):
        data = load(ws, stem)
        entry = _case(data, case)
        if not any(h["text"] == text for h in entry["hints"]):
            entry["hints"] = (entry["hints"] + [{"text": text, "at": time.strftime("%Y-%m-%d %H:%M:%S"), "session": session_id}])[-MAX_HINTS:]
        _save(ws, stem, data)


def remember_repair(ws, stem, case, note, session_id):
    """What the assistant found and changed after a failed run, so the next session does not start by repeating it."""
    with locked(_file(ws, stem)):
        data = load(ws, stem)
        entry = _case(data, case)
        entry["repairs"] = (entry["repairs"] + [{"cause": note["cause"], "change": note["change"], "at": time.strftime("%Y-%m-%d %H:%M:%S"), "session": session_id}])[-MAX_REPAIRS_KEPT:]
        _save(ws, stem, data)


def remember_result(ws, stem, case, status, reason, session_id):
    with locked(_file(ws, stem)):
        data = load(ws, stem)
        entry = _case(data, case)
        entry.update(status=status, reason=reason or "", session=session_id, at=time.strftime("%Y-%m-%d %H:%M:%S"))
        _save(ws, stem, data)


def recall(ws, stem, case):
    """What is known about `case` from earlier runs: {status, reason, session, hints[text]} or None."""
    entry = load(ws, stem)["cases"].get(case)
    if not entry or not (entry["status"] or entry["hints"] or entry.get("repairs")):
        return None
    return {"status": entry["status"], "reason": entry["reason"], "session": entry["session"], "hints": [h["text"] for h in entry["hints"]],
            "tried": [f"{r['cause']} -> {r['change']}" for r in entry.get("repairs", [])]}


def hints(ws, stem, case):
    return (load(ws, stem)["cases"].get(case) or {}).get("hints", [])


def summary(ws, stem, case_ids):
    """Lines for `session new`: what the earlier runs of this file already say."""
    cases = load(ws, stem)["cases"]
    failed = [c for c in case_ids if cases.get(c, {}).get("status") in ("FAIL", "BLOCKED", "PENDING")]
    hinted = [c for c in case_ids if cases.get(c, {}).get("hints")]
    lines = []
    if failed:
        lines.append("Did not pass last time: " + ", ".join(f"{c} ({cases[c]['status']})" for c in failed))
    if hinted:
        lines.append("The user already gave hints for: " + ", ".join(hinted))
    return lines


def backfill(ws, stem, data):
    """Read the sessions that already exist on this file into `data`, oldest first, so the latest word on each case wins."""
    for sid in session_ids(ws):
        state = read_json(session_dir(ws, sid) / "state.json", None)
        if not state or stem_of(state.get("file", "")) != stem:
            continue
        for case, items in (state.get("hints") or {}).items():
            entry = _case(data, case)
            for h in items:
                if not any(x["text"] == h["text"] for x in entry["hints"]):
                    entry["hints"].append({"text": h["text"], "at": h.get("at", ""), "session": sid})
            entry["hints"] = entry["hints"][-MAX_HINTS:]
        for case, result in (state.get("results") or {}).items():
            _case(data, case).update(status=result["status"], reason=result.get("reason", ""), session=sid, at=state.get("created", ""))
