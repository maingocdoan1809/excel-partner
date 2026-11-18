# -*- coding: utf-8 -*-
"""PROJECT HOOKS - the only project-specific data code of the autotest (qa.config.json data.hooks = "project.hooks").
Everything is optional: delete what you do not need; the framework reports what it could not verify.
Rules: READ-ONLY except CLEANERS; CLEANERS remove ONLY objects whose name starts with cfg.prefix (checked again below).
Credentials come from environment variables, never from this file."""
from qakit.core import guard
from qakit.core.config import cfg


def check_data():
    """Preflight: prove the framework reads the SAME data source the app under test uses (e.g. compare a name shown in the UI with the DB).
    Return a short text; start it with 'ERROR' when it does not match."""
    return "not implemented: integrity of the data source is not proven"


def snapshot():
    """Read-only dump of what a test run may touch: {collection: {id: object}}. Example with an HTTP API:
        import json, urllib.request
        items = json.load(urllib.request.urlopen(cfg.url("/api/items")))
        return {"items": {str(i["id"]): i for i in items}}
    Example with SQL (any DB-API driver):  rows = conn.execute("SELECT id, name, updated FROM users").fetchall()"""
    return {}


def label(collection, obj):
    """Short name of an object in the integrity diff (it must contain the visible name so test objects are recognised by the prefix)."""
    return str(obj.get("name") or obj.get("title") or obj.get("email") or obj.get("id")) if isinstance(obj, dict) else str(obj)


def discover():
    """Test objects that still exist (leftovers of earlier runs): [{"kind": "...", "name": "...", "id": ...}]."""
    return []


def _remove_item(rec):
    guard.check(guard.WRITE_TEST, rec["name"], "delete")      # never anything without the prefix
    raise NotImplementedError("remove the object through the UI page object or the public API, then return True")


CLEANERS = {
    # "item": _remove_item,
}
CLEANUP_ORDER = []     # kinds in removal order (dependencies first), e.g. ["order_line", "order", "customer"]
