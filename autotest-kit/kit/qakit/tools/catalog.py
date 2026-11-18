# -*- coding: utf-8 -*-
"""TEST SUITE CATALOG: one record per testcase ID across every source workbook + which suite automates it (read-only on sources).
  python -m qakit.tools.catalog build [--dry-run] [--note "why"]   # create/update test_suite/catalog.json (+ history/<build>.json, CHANGELOG.md)
  python -m qakit.tools.catalog validate                           # consistency/completeness/duplicates/automation problems (exit 1 on errors;
                                                                    #  a source workbook saved after the build = ERROR "catalog out of date")
  python -m qakit.tools.catalog show [--file 09] [--id USR-001]    # structure + status
  python -m qakit.tools.catalog diff [--base previous|<build id>]  # current catalog vs an older build
  python -m qakit.tools.catalog coverage                           # automated / not automated / deprecated per file and per タイプ
  python -m qakit.tools.catalog ack USR-010 USR-011 --note "..."   # clear needs_review after a human checked the automation
Catalog folder: test_suite/ (sandbox: `--sandbox <name>` or env QA_SUITE_DIR=test_suite_<name> -> an isolated catalog, e.g. for a demo). Sources: <folder>/sources.json {"sources": [globs relative to auto-test]}. Source workbooks are NEVER written.
Lifecycle per ID: new -> active -> changed (content hash differs; needs_review when automated) -> deprecated (missing from every source; kept,
never deleted; returns to active if it reappears). IDs, first_seen and history are preserved across builds."""
import argparse, datetime, glob, hashlib, importlib, inspect, json, os, re, sys, traceback
from pathlib import Path
from ..core.config import WORK, cfg
from ..reporting import spec as spec_mod

sys.stdout.reconfigure(encoding="utf-8")
DIR = CATALOG = HISTORY = CHANGELOG = SOURCES = None


def use(folder):
    """Point the module at a catalog folder under auto-test: 'test_suite' (real) or 'test_suite_<sandbox>' (e.g. test_suite_demo)."""
    global DIR, CATALOG, HISTORY, CHANGELOG, SOURCES
    DIR = WORK / folder
    CATALOG, HISTORY, CHANGELOG, SOURCES = DIR / "catalog.json", DIR / "history", DIR / "CHANGELOG.md", DIR / "sources.json"


def sandboxes():
    """Names of the sandbox catalogs that exist (folder test_suite_<name> with sources.json)."""
    return sorted(p.name[len("test_suite_"):] for p in WORK.glob("test_suite_*") if (p / "sources.json").exists())


use(os.environ.get("QA_SUITE_DIR", "test_suite"))       # env QA_SUITE_DIR or --sandbox <name> -> isolated catalog (demo / experiments)
DEFAULT_SOURCES = {"sources": list(cfg.raw["testcases"]["sources"]),
                   "_note": "globs relative to the autotest root; earlier entries are the PRIMARY source of an ID that appears in several files"}
HASHED = ("view", "role", "pre", "act", "exp", "cls", "typ")


def rel(p):
    try:
        return str(Path(p).resolve().relative_to(WORK)).replace("\\", "/")
    except ValueError:
        return str(p).replace("\\", "/")


def sources():
    globs = (json.load(open(SOURCES, encoding="utf-8")) if SOURCES.exists() else DEFAULT_SOURCES)["sources"]
    out = []
    for g in globs:
        out += [rel(f) for f in sorted(glob.glob(str(WORK / g))) if "~$" not in f and rel(f) not in out]
    return out


def fhash(path):
    """Byte hash of a source workbook (any save changes it; content-level changes are then classified by `build`)."""
    try:
        return hashlib.sha1((WORK / path).read_bytes()).hexdigest()[:16]
    except OSError:
        return None


def stale(cat):
    """Source workbooks that differ from what the catalog was built from -> {file: reason}. Empty = catalog is up to date.
    reason: 'changed' (saved since the build) | 'missing' (deleted / renamed) | 'new' (matches sources.json but not in the build) |
    'untracked' (catalog built before source hashes were recorded)."""
    known = cat.get("source_hashes")
    if known is None:
        return {f: "untracked" for f in cat.get("sources", [])}
    out = {}
    for f, h in known.items():
        now = fhash(f)
        if now is None:
            out[f] = "missing"
        elif now != h:
            out[f] = "changed"
    for f in sources():
        if f not in known:
            out[f] = "new"
    return out


STALE_TEXT = {"changed": "saved after the last build", "missing": "missing (deleted / renamed)", "new": "new file, not in the catalog yet",
              "untracked": "catalog built without source hashes"}


def chash(c):
    return hashlib.sha1(json.dumps([re.sub(r"\s+", " ", str(c.get(k, ""))).strip() for k in HASHED], ensure_ascii=False).encode()).hexdigest()[:12]


def spec_version(path):
    """'Phiên bản' / 'バージョン' / 'Version' value of the cover sheet, if any."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(WORK / path, read_only=True)
        for ws in wb.worksheets[:2]:
            for row in ws.iter_rows(min_row=1, max_row=20, max_col=8, values_only=True):
                cells = [x for x in row if x not in (None, "")]
                for i, x in enumerate(cells[:-1]):
                    if str(x).strip() in cfg.raw["testcases"]["version_labels"]:
                        return str(cells[i + 1])
    except Exception:
        pass
    return ""


def discover_suites():
    """Import every suites/<name>/cases.py of the project and read its registered case IDs. Import errors are returned, never hidden."""
    from ..core.tc import Suite
    suites, errors = {}, {}
    if str(WORK) not in sys.path:
        sys.path.insert(0, str(WORK))
    for d in sorted((WORK / "suites").iterdir()) if (WORK / "suites").is_dir() else []:
        if not (d / "cases.py").exists():
            continue
        try:
            m = importlib.import_module(f"suites.{d.name}.cases")
            s = getattr(m, "suite", None)
            if not isinstance(s, Suite):
                errors[d.name] = "cases.py has no `suite = Suite(...)`"
                continue
            fns = list(s.cases.values())
            matrix = bool(fns) and len(inspect.signature(fns[0]).parameters) >= 3
            todo = sum(p.read_text(encoding="utf-8").count("t.todo(") for p in d.glob("*.py"))
            sp = Path(s.spec_path) if s.spec_path else None
            if sp and not sp.is_absolute():
                sp = WORK / sp                                  # suites are run with cwd = autotest root
            suites[d.name] = dict(spec=rel(sp) if sp else "", spec_exists=bool(sp) and sp.exists(),
                                  mode="matrix" if matrix else "plain", ids=[str(k) for k in s.cases], todo=todo)
        except Exception as e:
            errors[d.name] = f"{type(e).__name__}: {str(e)[:200]}"
    return suites, errors


def load(path=None):
    path = Path(path or CATALOG)
    return json.load(open(path, encoding="utf-8")) if path.exists() else None


def build(dry=False, note=""):
    from ..core.tc import targets_of
    prev = load() or {"build": None, "cases": {}}
    bid = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    files = sources()
    suites, errors = discover_suites()
    by_spec = {}
    for name, s in suites.items():
        by_spec.setdefault(s["spec"], []).append(name)
    seen, cases, conflicts = {}, {}, []
    for f in files:
        try:
            rows = spec_mod.read(WORK / f)
        except Exception as e:
            errors[f] = f"cannot read source: {e}"
            continue
        ver = spec_version(f)
        for c in rows:
            tid = str(c["no"]).strip()
            h = chash(c)
            if tid in seen:                                     # same ID in a later source (e.g. a demo subset): primary stays first
                if seen[tid]["hash"] != h:
                    conflicts.append(f"{tid}: content differs between {seen[tid]['file']} and {f}")
                seen[tid]["also_in"].append(f)
                continue
            seen[tid] = dict(id=tid, title=str(c.get("view", "")).strip(), file=f, file_version=ver, sheet=c["_sheet"], row=c["_row"],
                             cls=str(c.get("cls", "")).strip(), typ=str(c.get("typ", "")).strip(), role=str(c.get("role", "")).strip(),
                             targets=[f"{e}:{r}" for e, r in targets_of(c.get("role"))],
                             steps=len(c.get("act_lines") or []), hash=h, also_in=[])
    log = {"new": [], "changed": [], "deprecated": [], "reactivated": [], "needs_review": []}
    for tid, e in seen.items():
        old = prev["cases"].get(tid)
        auto = [dict(suite=n, mode=suites[n]["mode"], primary=suites[n]["spec"] == e["file"]) for n in suites if tid in suites[n]["ids"]]
        auto.sort(key=lambda a: (not a["primary"], a["suite"]))
        e["automation"] = auto
        if not old:
            e.update(status="new" if prev["build"] else "active", first_seen=bid, changed_in=bid, needs_review=False, history=[])
            if prev["build"]:
                log["new"].append(tid)
        else:
            e.update(first_seen=old["first_seen"], history=old.get("history", []), needs_review=old.get("needs_review", False), changed_in=old.get("changed_in"))
            if old.get("status") == "deprecated":
                e.update(status="active", changed_in=bid)
                e["history"].append(dict(build=bid, event="reactivated"))
                log["reactivated"].append(tid)
            elif old["hash"] != e["hash"]:
                e.update(status="changed", changed_in=bid)
                e["history"].append(dict(build=bid, event="changed", old_hash=old["hash"], new_hash=e["hash"],
                                         fields=[k for k in ("title", "role", "cls", "typ", "steps") if str(old.get(k)) != str(e.get(k))] or ["pre/act/exp"]))
                if auto:
                    e["needs_review"] = True                    # automation was written against the old text
                log["changed"].append(tid)
            else:                                               # unchanged: new -> active; changed stays changed until reviewed (ack)
                e["status"] = "changed" if old.get("status") == "changed" and e["needs_review"] else "active"
        if e["needs_review"]:
            log["needs_review"].append(tid)
        cases[tid] = e
    for tid, old in prev["cases"].items():                     # never delete: keep missing IDs as deprecated
        if tid not in cases:
            o = dict(old)
            if o.get("status") != "deprecated":
                o.update(status="deprecated", changed_in=bid)
                o.setdefault("history", []).append(dict(build=bid, event="deprecated", reason="ID not found in any source"))
                log["deprecated"].append(tid)
            o["automation"] = [dict(suite=n, mode=suites[n]["mode"], primary=False) for n in suites if tid in suites[n]["ids"]]
            cases[tid] = o
    cat = dict(build=bid, previous_build=prev["build"], built_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), sources=files,
               source_hashes={f: fhash(f) for f in files}, suites=suites, suite_errors=errors, conflicts=conflicts, cases=dict(sorted(cases.items())))
    summary = {k: len(v) for k, v in log.items()}
    print(f"build {bid}: {len(cases)} IDs from {len(files)} sources, {len(suites)} suites ({len(errors)} load errors) | " + ", ".join(f"{k}={v}" for k, v in summary.items()))
    for k in ("new", "changed", "deprecated", "reactivated", "needs_review"):
        if log[k]:
            print(f"  {k}: {', '.join(log[k][:60])}{' …' if len(log[k]) > 60 else ''}")
    for c in conflicts[:20]:
        print("  conflict:", c)
    if dry:
        print("(dry-run: nothing written)")
        return cat, log
    HISTORY.mkdir(parents=True, exist_ok=True)
    if not SOURCES.exists():
        SOURCES.write_text(json.dumps(DEFAULT_SOURCES, ensure_ascii=False, indent=1), encoding="utf-8")
    txt = json.dumps(cat, ensure_ascii=False, indent=1)
    CATALOG.write_text(txt, encoding="utf-8")
    (HISTORY / f"{bid}.json").write_text(txt, encoding="utf-8")
    entry = [f"## {bid}" + (f" — {note}" if note else ""), "",
             f"- Sources: {len(files)} file(s) · {len(cases)} ID · suites: {len(suites)} (load errors: {', '.join(errors) or 'none'})",
             *[f"- {k}: {len(v)}" + (f" ({', '.join(v[:30])}{' …' if len(v) > 30 else ''})" if v else "") for k, v in log.items()], ""]
    old = CHANGELOG.read_text(encoding="utf-8") if CHANGELOG.exists() else "# Test Suite change log (written by qakit.tools.catalog)\n\n"
    head, _, rest = old.partition("\n\n")
    CHANGELOG.write_text(head + "\n\n" + "\n".join(entry) + "\n" + rest, encoding="utf-8")
    print("written", rel(CATALOG), "+", rel(HISTORY / f"{bid}.json"))
    return cat, log


def need():
    cat = load()
    if not cat:
        sys.exit("no catalog yet: run `python -m qakit.tools.catalog build`")
    return cat


def validate():
    from ..core.tc import recognised
    cat = need()
    err, warn = [], []
    for f, e in cat["suite_errors"].items():
        err.append(f"suite/source load error: {f}: {e}")
    for f, why in stale(cat).items():
        err.append(f"catalog out of date: {f}: {STALE_TEXT[why]} -> run `/system-test update` (qakit.tools.catalog build) first")
    for c in cat["conflicts"]:
        warn.append(f"same ID, different content: {c}")
    for name, s in cat["suites"].items():
        if s["spec"] and not s["spec_exists"]:
            err.append(f"suite {name}: spec file not found: {s['spec']}")
        if s["spec"] and re.match(r"^[A-Za-z]:/", s["spec"]):
            warn.append(f"suite {name}: spec outside auto-test ({s['spec']}) -> move the workbook next to the suite or into a source folder")
        if not s["ids"]:
            warn.append(f"suite {name}: registers no case (script-style) -> cannot be selected by ID, only run whole (suite={name})")
        if s["todo"]:
            warn.append(f"suite {name}: {s['todo']} t.todo() step(s) left (those steps report NOT EXECUTED)")
        if s["spec"] in cat["sources"]:
            orphan = [i for i in s["ids"] if i not in cat["cases"]]
            if orphan:
                err.append(f"suite {name}: IDs not in its spec/catalog: {', '.join(orphan[:20])}")
    for tid, c in cat["cases"].items():
        if c["status"] == "deprecated" and c.get("automation"):
            warn.append(f"{tid}: deprecated but still automated by {', '.join(a['suite'] for a in c['automation'])}")
        if c.get("needs_review"):
            warn.append(f"{tid}: spec changed after automation was written ({', '.join(a['suite'] for a in c['automation'])}) -> review + `catalog ack`")
        if c["status"] == "deprecated":
            continue
        if not c.get("steps"):
            err.append(f"{tid}: no action steps ({c['file']} row {c['row']})")
        M = cfg.raw["matrix"]
        if (M.get("envs") or M.get("roles")) and c["role"] and recognised(c["role"]) == ([], []):
            warn.append(f"{tid}: role/env column not recognised ('{c['role'][:40]}') -> runner uses the defaults (add keywords to matrix.envs / matrix.roles)")
        if len([a for a in c.get("automation", []) if a["primary"]]) > 1:
            warn.append(f"{tid}: automated by several primary suites")
    print(f"catalog {cat['build']}: {len(err)} error(s), {len(warn)} warning(s)")
    for x in err:
        print("ERROR  ", x)
    for x in warn:
        print("WARN   ", x)
    return 1 if err else 0


def show(file=None, tid=None):
    cat = need()
    cs = [c for c in cat["cases"].values() if (not file or file in c["file"] or f"_{file}_" in c["file"]) and (not tid or c["id"] in tid.split(","))]
    if tid:
        for c in cs:
            print(json.dumps(c, ensure_ascii=False, indent=1))
        return
    print(f"catalog build {cat['build']} ({cat['built_at']}) · {len(cat['cases'])} IDs · sources {len(cat['sources'])}")
    by = {}
    for c in cs:
        by.setdefault(c["file"], []).append(c)
    print("| File | Version | IDs | active | new | changed | deprecated | automated | needs review |\n|---|---|---|---|---|---|---|---|---|")
    for f, L in sorted(by.items()):
        n = lambda s: sum(1 for c in L if c["status"] == s)
        print(f"| {Path(f).name} | {L[0].get('file_version', '')} | {len(L)} | {n('active')} | {n('new')} | {n('changed')} | {n('deprecated')} | "
              f"{sum(1 for c in L if c.get('automation'))} | {sum(1 for c in L if c.get('needs_review'))} |")
    print("\nSuites:")
    for name, s in cat["suites"].items():
        miss = " [MISSING]" if s["spec"] and not s["spec_exists"] else ""
        todo = f", todo={s['todo']}" if s["todo"] else ""
        print(f"- {name}: {len(s['ids'])} case(s), mode={s['mode']}, spec={s['spec'] or '(none)'}{miss}{todo}")
    for name, e in cat["suite_errors"].items():
        print(f"- {name}: LOAD ERROR {e}")
    if file:
        print()
        for c in cs:
            print(f"{c['id']}\t{c['status']}\t{','.join(a['suite'] for a in c.get('automation', [])) or '-'}\t{c['typ']}\t{c['cls']}\t{c['title'][:50]}")


def diff(base="previous"):
    cat = need()
    bid = cat["previous_build"] if base == "previous" else base
    if not bid:
        sys.exit("no previous build to compare with")
    old = load(HISTORY / f"{bid}.json")
    if not old:
        sys.exit(f"build {bid} not found in {rel(HISTORY)}")
    A, B = old["cases"], cat["cases"]
    added = [i for i in B if i not in A]
    removed = [i for i in A if i not in B]
    changed = [i for i in B if i in A and A[i]["hash"] != B[i]["hash"]]
    status = [f"{i}: {A[i]['status']} -> {B[i]['status']}" for i in B if i in A and A[i]["status"] != B[i]["status"]]
    auto = [i for i in B if i in A and [a["suite"] for a in A[i].get("automation", [])] != [a["suite"] for a in B[i].get("automation", [])]]
    print(f"diff {bid} -> {cat['build']}: +{len(added)} -{len(removed)} ~{len(changed)} status:{len(status)} automation:{len(auto)}")
    for t, L in (("added", added), ("removed (should not happen: IDs are kept)", removed), ("content changed", changed), ("status", status), ("automation changed", auto)):
        if L:
            print(f"  {t}: {', '.join(L[:80])}")


def coverage():
    cat = need()
    live = [c for c in cat["cases"].values() if c["status"] != "deprecated"]
    print(f"coverage (catalog {cat['build']}): {sum(1 for c in live if c.get('automation'))}/{len(live)} active IDs automated "
          f"· deprecated {len(cat['cases']) - len(live)}")
    for key, title in (("file", "File"), ("typ", "Type")):
        g = {}
        for c in live:
            g.setdefault(Path(c[key]).name if key == "file" else (c[key] or "(empty)"), []).append(c)
        print(f"\n| {title} | IDs | automated | not automated | % |\n|---|---|---|---|---|")
        for k, L in sorted(g.items()):
            a = sum(1 for c in L if c.get("automation"))
            print(f"| {k} | {len(L)} | {a} | {len(L) - a} | {100 * a // max(len(L), 1)}% |")


def ack(ids, note):
    cat = need()
    for tid in ids:
        c = cat["cases"].get(tid)
        if not c:
            print("unknown", tid)
            continue
        c["needs_review"] = False
        if c["status"] == "changed":
            c["status"] = "active"
        c.setdefault("history", []).append(dict(build=cat["build"], event="reviewed", note=note))
    CATALOG.write_text(json.dumps(cat, ensure_ascii=False, indent=1), encoding="utf-8")
    old = CHANGELOG.read_text(encoding="utf-8") if CHANGELOG.exists() else "# Test Suite change log (written by qakit.tools.catalog)\n\n"
    head, _, rest = old.partition("\n\n")
    CHANGELOG.write_text(head + "\n\n" + f"## {datetime.datetime.now():%Y%m%d_%H%M%S} — review\n\n- reviewed: {', '.join(ids)}" + (f" — {note}" if note else "") + "\n\n" + rest, encoding="utf-8")
    print("acknowledged", ids)


def main():
    if "--sandbox" in sys.argv:                          # --sandbox <name> anywhere = catalog test_suite_<name>
        j = sys.argv.index("--sandbox")
        use("test_suite_" + sys.argv[j + 1])
        del sys.argv[j:j + 2]
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest="cmd", required=True)
    b = sp.add_parser("build"); b.add_argument("--dry-run", action="store_true"); b.add_argument("--note", default="")
    sp.add_parser("validate")
    s = sp.add_parser("show"); s.add_argument("--file"); s.add_argument("--id")
    d = sp.add_parser("diff"); d.add_argument("--base", default="previous")
    sp.add_parser("coverage")
    k = sp.add_parser("ack"); k.add_argument("ids", nargs="+"); k.add_argument("--note", default="")
    a = ap.parse_args()
    if a.cmd == "build":
        build(a.dry_run, a.note)
    elif a.cmd == "validate":
        return validate()
    elif a.cmd == "show":
        show(a.file, a.id)
    elif a.cmd == "diff":
        diff(a.base)
    elif a.cmd == "coverage":
        coverage()
    elif a.cmd == "ack":
        ack(a.ids, a.note)
    return 0


if __name__ == "__main__":
    sys.exit(main())
