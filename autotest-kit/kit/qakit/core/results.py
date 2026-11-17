# -*- coding: utf-8 -*-
"""results_<suite>.json (one record per testcase ID) - the single source for Excel builders, compare and consolidate."""
import json, os
from pathlib import Path
from .config import cfg


def path():
    d = Path(os.environ["QA_RESULTS_DIR"]) if os.environ.get("QA_RESULTS_DIR") else cfg.data_dir
    return d / (f"results_{cfg.suite}.json" if cfg.suite else "results.json")


def load():
    return json.load(open(path(), encoding="utf-8")) if path().exists() else {}


def record(tc, **kw):
    cfg.ensure()
    r = load()
    cur = r.get(str(tc), {})
    cur.update(kw)
    r[str(tc)] = cur
    json.dump(r, open(path(), "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def locate(p):
    """Image path stored in a results file -> an existing file (the output folder may have been moved)."""
    if not p or os.path.exists(p):
        return p
    parts = Path(p).parts
    for i in range(len(parts) - 1, -1, -1):
        if parts[i].lower() == "evidence":
            rest = parts[i + 1:]
            rest = rest[1:] if rest and rest[0] == "base" else rest
            q = cfg.evidence_root.joinpath(*rest)
            return str(q) if q.exists() else p
    return p
