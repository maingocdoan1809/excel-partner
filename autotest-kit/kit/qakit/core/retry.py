# -*- coding: utf-8 -*-
"""Retry that READS the screen between attempts. Never replay the same action blind.
    from qakit.core.retry import attempt
    attempt("create rule", lambda: page_obj.create_rule(name), fix=lambda d: page_obj.fill_missing(d["fields"]), pg=t.pg)
Attempt 1 fails -> diagnose (screenshot + validation messages + invalid fields) -> fix(diagnosis) repairs the CAUSE -> attempt 2.
Out of attempts, or the same failure as before (nothing was fixed) -> Diagnosed is raised carrying the diagnosis, so the case is recorded with the
real reason ("Customer Name required and empty") instead of a timeout. Budget: qa.config.json run.max_attempts (default 2)."""
import re
from . import browser, diagnose, events
from .config import cfg


class Refused(RuntimeError):
    """The screen did not accept what was done (a save refused, a record that is still missing)."""


class Diagnosed(RuntimeError):
    def __init__(self, label, error, diagnosis):
        super().__init__(f"{label}: {str(error)[:160]} | screen says: {diagnosis.get('summary', '')}")
        self.label, self.error, self.diagnosis = label, error, diagnosis


def budget():
    try:
        return max(1, min(5, int(cfg.raw.get("run", {}).get("max_attempts", 2))))
    except (TypeError, ValueError):
        return 2


def _sig(error, diagnosis):
    return re.sub(r"\d+", "#", f"{type(error).__name__}|{diagnosis.get('summary', '')}")[:160]


def attempt(label, fn, fix=None, tries=None, pg=None, case=None):
    tries = tries or budget()
    last = None
    for i in range(1, tries + 1):
        try:
            return fn()
        except Exception as e:                                       # noqa: BLE001 - every failure is diagnosed
            d = diagnose.snapshot(pg or browser.page(), f"{label}-try{i}", error=e, case=case)
            events.emit("phase", phase="fix", text=f"{label}: attempt {i}/{tries} failed")
            sig = _sig(e, d)
            if i == tries or sig == last or not fix:
                raise Diagnosed(label, e, d) from e
            last = sig
            fix(d)                                                    # repair the cause shown on the screen, then try again
