# -*- coding: utf-8 -*-
"""DATA-SAFETY GUARD (rule G2 of the system-test skill: never modify existing data; create / edit / delete only OWN test objects).
Every page-object method that changes something declares its risk and the target; the guard refuses when the target is not a test object.

Risk levels
  READ           viewing / listing / opening dialogs without saving                      -> always allowed
  WRITE_TEST     create / edit / delete an object whose name starts with the test prefix  -> allowed
  WRITE_EXISTING edit / delete an object WITHOUT the test prefix                          -> REFUSED (GuardError)
  GENERATE       actions that write as a side effect or cost money (AI generation, sending chat, evaluations) -> only on test objects;
                 bulk actions need the expected count (generate_scope)
  EXTERNAL       creates resources outside the app (cloud resources, real e-mails, payments, tickets) -> needs QA_ALLOW_EXTERNAL=1
Created test objects are recorded in safety/created_registry.json so tools.cleanup knows what to remove (dependencies first)."""
import datetime, json, os
from .config import cfg

READ, WRITE_TEST, WRITE_EXISTING, GENERATE, EXTERNAL = "READ", "WRITE_TEST", "WRITE_EXISTING", "GENERATE", "EXTERNAL"


class GuardError(RuntimeError):
    pass


def is_test(name):
    return str(name or "").strip().startswith(cfg.prefix)


def check(risk, target=None, action=""):
    """Raise GuardError if the action is not allowed. `target` = visible name of the object being changed."""
    if risk == READ:
        return
    if risk == EXTERNAL and os.environ.get("QA_ALLOW_EXTERNAL") != "1":
        raise GuardError(f"EXTERNAL action '{action}' on '{target}' needs the user's explicit OK (QA_ALLOW_EXTERNAL=1 after approval)")
    if risk in (WRITE_TEST, GENERATE, EXTERNAL) and not is_test(target):
        raise GuardError(f"REFUSED: '{action}' on '{target}' - not a test object (name must start with {cfg.prefix}). Existing data must not be modified.")
    if risk == WRITE_EXISTING:
        raise GuardError(f"REFUSED: '{action}' would modify existing data '{target}'")


def generate_scope(expected, actual, what="items"):
    """Bulk-action safety stop: the number of objects about to be processed must equal the expected number."""
    if int(actual) != int(expected):
        raise GuardError(f"SAFETY STOP: {actual} {what} selected, expected {expected}")


def _reg():
    return cfg.safety_dir / "created_registry.json"


def _load():
    return json.load(open(_reg(), encoding="utf-8")) if _reg().exists() else []


def track(kind, name, obj_id=None, parent=None, **extra):
    """Record a test object created through the UI (kind = free text agreed with project/hooks.py CLEANERS, e.g. user / item / order)."""
    cfg.ensure()
    reg = _load()
    reg.append(dict(kind=kind, name=name, id=obj_id, parent=parent, created_at=datetime.datetime.now().isoformat(timespec="seconds"), **extra))
    json.dump(reg, open(_reg(), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def untrack(kind, name):
    reg = [r for r in _load() if not (r["kind"] == kind and r["name"] == name)]
    json.dump(reg, open(_reg(), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def tracked(kind=None):
    return [r for r in _load() if kind in (None, r["kind"])]
