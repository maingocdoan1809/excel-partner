# -*- coding: utf-8 -*-
"""READ-ONLY data-integrity check: snapshot the data a test run may touch, then diff before / after.
  python -m qakit.reporting.snapshot take before        # before the first write  -> safety/snapshots/before.json
  python -m qakit.reporting.snapshot take after         # after the run + cleanup
  python -m qakit.reporting.snapshot diff before after  # every created / deleted / changed object; non-test objects = VIOLATION
Source (qa.config.json data.snapshot): "hooks" = project hooks snapshot(); "command" = data.command prints {collection: {id: object}} JSON;
"none" = no data check (the report says the integrity of the data was NOT verified)."""
import hashlib, json, os, subprocess, sys
from ..core import hooks
from ..core.config import cfg, WORK

D = cfg.raw["data"]


def digest(d):
    return hashlib.sha1(json.dumps(d, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()[:12]


def _label(coll, o):
    fn = getattr(hooks.module(), "label", None)
    if fn:
        return str(fn(coll, o))[:60]
    if isinstance(o, dict):
        return str(o.get("name") or o.get("title") or o.get("email") or o.get("id") or "")[:60]
    return str(o)[:60]


def collect():
    mode = D.get("snapshot", "none")
    if mode == "hooks":
        fn = getattr(hooks.module(), "snapshot", None)
        if not fn:
            raise RuntimeError("data.snapshot = hooks but project hooks have no snapshot()")
        return fn()
    if mode == "command":
        p = subprocess.run(D["command"], shell=True, cwd=WORK, capture_output=True, text=True, encoding="utf-8")
        if p.returncode:
            raise RuntimeError(f"data.command failed: {p.stderr[-300:]}")
        return json.loads(p.stdout)
    return None


def take(name):
    data = collect()
    if data is None:
        print("data.snapshot = none: integrity of the data is NOT verified")
        return None
    ign = set(D.get("ignore_fields", ["updated_at", "updatedAt", "last_login", "lastLogin", "_ts", "_etag"]))
    out = {c: {str(i): {"hash": digest({k: v for k, v in o.items() if k not in ign} if isinstance(o, dict) else o), "label": _label(c, o)}
               for i, o in objs.items()} for c, objs in data.items()}
    cfg.snapshots_dir.mkdir(parents=True, exist_ok=True)
    json.dump(out, open(cfg.snapshots_dir / f"{name}.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print({c: len(v) for c, v in out.items()})
    return out


def diff(a, b):
    pa, pb = cfg.snapshots_dir / f"{a}.json", cfg.snapshots_dir / f"{b}.json"
    if not (pa.exists() and pb.exists()):
        print("no snapshots: integrity of the data is NOT verified")
        return -1
    A, B = json.load(open(pa, encoding="utf-8")), json.load(open(pb, encoding="utf-8"))
    bad = n = 0
    for c in sorted(set(A) | set(B)):
        x, y = A.get(c, {}), B.get(c, {})
        rows = [("DELETED", i, x[i]["label"]) for i in x.keys() - y.keys()] + [("CREATED", i, y[i]["label"]) for i in y.keys() - x.keys()] + \
               [("CHANGED", i, y[i]["label"]) for i in x.keys() & y.keys() if x[i]["hash"] != y[i]["hash"]]
        for kind, i, lab in rows:
            n += 1
            test = lab.startswith(cfg.prefix)
            bad += not test
            print(f"{kind:8} {c}/{i} {lab}" + ("" if test else "   <-- NOT a test object: VIOLATION"))
    print("no differences" if not n else f"{n} difference(s), {bad} violation(s)")
    return bad


if __name__ == "__main__":
    if sys.argv[1:2] == ["take"]:
        take(sys.argv[2])
    elif sys.argv[1:2] == ["diff"]:
        n = diff(sys.argv[2], sys.argv[3])
        sys.exit(3 if n < 0 else 1 if n > 0 else 0)          # 0 verified clean, 1 violation, 3 UNVERIFIED (snapshot missing)
    else:
        print(__doc__)
