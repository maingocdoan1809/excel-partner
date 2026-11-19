"""events.jsonl: one JSON line per event, appended as things happen. The dashboard, the Excel report and a person tailing the file all read this and nothing else.

Kinds and their required fields (every event also has `ts` and `kind`):
  session_started  workspace, file
  phase            phase, text                      the phase just entered / what the assistant is doing now
  ask              question [case]                  the assistant is waiting for the user
  answered         answer [case, files]               files = the paths written in the answer, now references
  ref_added        path, by [name, what (file|folder), case]        the user shared a file or folder by its path (nothing is uploaded; the assistant reads it where it is)
  confirmed        by, strategy_sha
  stopped / resumed  by                                  the user stopped the test / let it go on
  hint             case, text                            the user's advice for a case; it goes back in the queue with a fresh budget
  case_planned     case, title, steps[{n,action,expected}], order
  case_started     case, attempt
  step             case, attempt, n, status, caption, expected, actual, images[{raw,annotated}]
                   PASS and FAIL steps MUST carry at least one image pair: a step without a real picture is not recorded.
  attempt_finished case, attempt, status PASS|FAIL, reason [failed_step, script_sha]     one run of a case; the case may be repaired and run again
  repair           case, attempt, cause, evidence, change [reanalysis, probed, script_sha, timeout_s, unchanged_reason]
                   the assistant's note after a failed run: what it found, how it knows, what it changed, how it now reads the case. A rerun needs one.
  try_started      case, step [mode]                                           the assistant is running code for a step on the real page RIGHT NOW (closed by the `explore` event that follows)
  built            case, steps                                                 the script was built from the verified steps: the case waits to be run
  probe            url, count [case, filter]                                  the assistant looked at the real page (oqa probe)
  explore          case, step, ok [code, url, error]                           the assistant tried code for a step on the real page (oqa try)
  case_finished    case, status, attempts [reason]                           the case's final result
  session_finished counts
"""
import json
import os
import time
from pathlib import Path

from .store import SchemaError

KINDS = {
    "session_started": ("workspace", "file"),
    "phase": ("phase", "text"),
    "ask": ("question",),
    "answered": ("answer",),
    "ref_added": ("path", "by"),
    "confirmed": ("by", "strategy_sha"),
    "stopped": ("by",),
    "resumed": ("by",),
    "hint": ("case", "text"),
    "case_planned": ("case", "title", "steps", "order"),
    "case_started": ("case", "attempt"),
    "step": ("case", "attempt", "n", "status", "caption", "expected", "actual", "images"),
    "attempt_finished": ("case", "attempt", "status", "reason"),
    "repair": ("case", "attempt", "cause", "evidence", "change"),
    "try_started": ("case", "step"),
    "built": ("case", "steps"),
    "probe": ("url", "count"),
    "explore": ("case", "step", "ok"),
    "case_finished": ("case", "status", "attempts"),
    "session_finished": ("counts",),
}
USER_KINDS = ("confirmed", "answered", "ref_added", "hint", "stopped", "resumed")        # what only the user does: the assistant waits for these
STEP_STATUSES = ("PASS", "FAIL", "SKIPPED")
CASE_STATUSES = ("PASS", "FAIL", "PENDING", "BLOCKED", "NOT EXECUTED")


def validate(kind, data):
    if kind not in KINDS:
        raise SchemaError(f"unknown event kind '{kind}'", "event")
    problems = [f"'{kind}' needs '{key}'" for key in KINDS[kind] if key not in data]
    if kind == "step" and not problems:
        if data["status"] not in STEP_STATUSES:
            problems.append(f"step status must be one of {', '.join(STEP_STATUSES)}")
        elif data["status"] != "SKIPPED":
            images = data["images"]
            if not isinstance(images, list) or not images or not all(isinstance(i, dict) and i.get("raw") and i.get("annotated") for i in images):
                problems.append("a PASS or FAIL step needs at least one picture, each with 'raw' and 'annotated' (the red-box one)")
    if kind == "case_finished" and not problems and data["status"] not in CASE_STATUSES:
        problems.append(f"case status must be one of {', '.join(CASE_STATUSES)}")
    if problems:
        raise SchemaError(problems, "event")


def append(session_dir, kind, **data):
    """Validate and append one event. Unlike progress reporting in the old kit, a malformed event is an error: the file is the evidence."""
    validate(kind, data)
    record = {"ts": time.strftime("%Y-%m-%d %H:%M:%S")}
    record.update(data)
    record["kind"] = kind
    file = Path(session_dir) / "events.jsonl"
    file.parent.mkdir(parents=True, exist_ok=True)
    line = (json.dumps(record, ensure_ascii=False, default=str) + "\n").encode("utf-8")
    fd = os.open(file, os.O_WRONLY | os.O_APPEND | os.O_CREAT)
    try:
        os.write(fd, line)
    finally:
        os.close(fd)
    return record


def read(session_dir):
    """All events; a line still being written is skipped."""
    file = Path(session_dir) / "events.jsonl"
    out = []
    if file.exists():
        for line in file.read_text(encoding="utf-8").splitlines():
            try:
                out.append(json.loads(line))
            except ValueError:
                continue
    return out
