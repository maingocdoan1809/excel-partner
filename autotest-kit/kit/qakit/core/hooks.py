# -*- coding: utf-8 -*-
"""Project hooks: the ONLY place where project-specific data access lives (data.hooks in qa.config.json, default project/hooks.py).
All functions are optional; the framework degrades gracefully (and says so in the report) when one is missing.

  snapshot() -> {collection: {id: object}}       read-only dump of the data a test run may touch (DB / API). Used for the before / after
                                                 integrity diff: anything created / deleted / changed that is not a test object = violation.
  label(collection, obj) -> str                  short name of an object in the diff report (default: name / title / email / id)
  discover() -> [ {kind, name, id, ...} ]        test objects (name starts with the test prefix) that still exist -> cleanup plan
  CLEANERS = {kind: fn(record) -> bool}          remove one test object (prefer the UI / public API; True = removed). Order of kinds in
                                                 CLEANUP_ORDER (dependencies first).
  check_data() -> str                            preflight: prove the framework reads the SAME database / API the app under test uses
  backup_roles(folder) -> path                   save whatever the role switching changes, before the first switch
  data_command = "..."                           (alternative to snapshot()) command printing the snapshot JSON on stdout"""
import importlib, sys
from functools import lru_cache
from .config import cfg, WORK


@lru_cache(maxsize=None)
def module():
    if str(WORK) not in sys.path:
        sys.path.insert(0, str(WORK))
    name = cfg.raw["data"].get("hooks") or "project.hooks"
    try:
        return importlib.import_module(name)
    except ModuleNotFoundError:
        return object()          # no hooks: every optional function is missing
