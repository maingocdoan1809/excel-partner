# -*- coding: utf-8 -*-
"""EXECUTION HISTORY: one folder per run (test_execution_output/runs/<run_id>/ = QA_OUT of that run) + runs/index.json. Nothing is overwritten.
run_id = <YYYYMMDD_HHMMSS>_<env>[_<label>]; folder layout: test_execution_output/README.md (SUMMARY.md first).
  python -m qakit.tools.runs new --plan <plan.json> [--label x]    # create the run folder, copy the plan in (prints run_id + QA_OUT)
  python -m qakit.tools.runs finish <run|latest> [--status cancelled]   # aggregate results_<suite>.json of the run into the index
  python -m qakit.tools.runs summary <run|latest>                  # (re)write SUMMARY.md of the run (finish does it too)
  python -m qakit.tools.runs list                                  # runs (new layout) + legacy outputs (read-only)
  python -m qakit.tools.runs show <run|latest>                     # per-ID status + failure-cause HINT (the agent confirms it from the evidence)
  python -m qakit.tools.runs compare <run|latest> [--base previous|<run>]   # new FAIL / fixed / still failing / changed / unchanged
  python -m qakit.tools.runs consolidate <run|latest>              # partial run -> excel/<spec>_evidence_<ja|vi>_tong_hop.xlsx: EVERY case of the
                                                                    #  current spec, latest result (this run, else the newest earlier run of the suite)
Legacy (before 2026-11-03, read-only): test_execution_output/legacy/flat/ = run 'legacy'; legacy/<dir>/ = 'legacy:<dir>'.
`--base previous` = for every ID the latest earlier result in any run (legacy counts as oldest)."""
import argparse, datetime, json, os, re, sys
from pathlib import Path
from ..core.config import WORK, cfg

sys.stdout.reconfigure(encoding="utf-8")
OUT = cfg.output_root
RUNS = OUT / "runs"
LEGACY = OUT / "legacy"
INDEX = RUNS / "index.json"
RUN_DIRS = ("excel", "evidence/base", "reports", "data", "safety", "logs")
ORDER = ("FAIL", "NOT EXECUTED", "PENDING", "BLOCKED", "PASS")


def index():
    return json.load(open(INDEX, encoding="utf-8")) if INDEX.exists() else []


def save_index(ix):
    RUNS.mkdir(parents=True, exist_ok=True)
    INDEX.write_text(json.dumps(ix, ensure_ascii=False, indent=1), encoding="utf-8")


def result_files(d):
    """results_<suite>.json of a run folder (data/; logs/ in the old flat layout)."""
    return sorted((d / "data").glob("results_*.json")) or sorted((d / "logs").glob("results_*.json"))


def legacy():
    """[(run_id, folder, mtime)] for outputs written before per-run folders existed (oldest first)."""
    out = []
    for d in sorted(LEGACY.iterdir()) if LEGACY.exists() else []:
        if d.is_dir() and result_files(d):
            out.append(("legacy" if d.name == "flat" else f"legacy:{d.name}", d, max(p.stat().st_mtime for p in result_files(d))))
    return sorted(out, key=lambda x: x[2])


def resolve(run):
    """run id (or its folder path) | 'latest' -> (run_id, folder)."""
    run = Path(run).name if ("/" in run or "\\" in run) else run
    ix = index()
    if run == "latest":
        if not ix:
            sys.exit("no run in runs/index.json yet")
        run = ix[-1]["run_id"]
    for rid, d, _ in legacy():
        if rid == run:
            return rid, d
    if (RUNS / run).is_dir():
        return run, RUNS / run
    sys.exit(f"unknown run: {run}")


def results_of(run):
    """{(suite, id): record} of a run (results_<suite>.json files)."""
    _, d = resolve(run)
    out = {}
    for p in result_files(d):
        suite = p.stem[len("results_"):]
        for k, v in json.load(open(p, encoding="utf-8")).items():
            out[(suite, str(k))] = v
    return out


def hint(r):
    """Deterministic first guess of the failure cause; the agent must confirm it from screenshots / log (product vs automation vs env vs data)."""
    st = r.get("status", "")
    rem = (r.get("remark") or "") + " " + " ".join(str(s.get("remark") or s.get("actual") or "") for s in r.get("steps", [])[-3:])
    if st == "PASS":
        return ""
    if re.search(r"script / runner error|Traceback|TimeoutError|evidence capture failed|TODO|not written yet|selector not found", rem):
        return "automation (script / runner)"
    if st == "FAIL":
        return "product (candidate: confirm with screenshot + source)"
    if re.search(r"login|MFA|SSO|switching not configured|cannot switch|permission|NEEDS", rem, re.I):
        return "environment / permission (needs a human or configuration)"
    if st == "PENDING":
        return "needs a human decision / action"
    if re.search(r"data|account|precondition|quota", rem, re.I):
        return "test data / precondition"
    if st == "BLOCKED":
        return "environment / observability"
    return "automation (not executed)" if st == "NOT EXECUTED" else "unknown"


def new(plan_path, label=""):
    plan = json.load(open(plan_path, encoding="utf-8"))
    if plan.get("stale"):                       # tools.scope found source workbooks saved after the catalog build
        sys.exit("plan is based on an out-of-date catalog (" + ", ".join(plan["stale"]) + "): run `catalog build` and `scope` again")
    env = "custom" if os.environ.get("QA_BASE_URL") else re.sub(r"[^A-Za-z0-9_-]+", "-", cfg.name)   # QA_BASE_URL: another slot / environment
    rid = "_".join([datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), env] + ([re.sub(r"[^A-Za-z0-9_-]+", "-", label)] if label else []))
    d = RUNS / rid
    for sub in RUN_DIRS:
        (d / sub).mkdir(parents=True, exist_ok=True)
    plan.update(run_id=rid, env=env, out=str(d))
    (d / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")   # the source plan file is left untouched
    try:
        from ..core import events
        events.emit_to(d / "events.jsonl", "run_planned", run_id=rid, label=label, cases=events.describe_plan(plan),
                       not_automated=plan.get("not_automated", []))
    except Exception:
        pass
    ix = index()
    ix.append(dict(run_id=rid, env=env, label=label, created=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), scope=plan.get("scope", []),
                   suites=[i["suite"] for i in plan.get("items", [])], ids=sum(len(i["ids"]) for i in plan.get("items", [])),
                   not_automated=len(plan.get("not_automated", [])), status="created"))
    save_index(ix)
    print("run_id", rid)
    print("QA_OUT", d)
    print("plan", d / "plan.json")
    return rid, d


def finish(run, status="finished"):
    rid, d = resolve(run)
    res = results_of(rid)
    cnt = {}
    for r in res.values():
        cnt[r.get("status", "?")] = cnt.get(r.get("status", "?"), 0) + 1
    plan_file = d / "plan.json"
    planned = set()
    if plan_file.exists():
        for it in json.load(open(plan_file, encoding="utf-8")).get("items", []):
            planned |= {(it["suite"], str(i)) for i in it["ids"]}
    missing = sorted(planned - {(s, str(i)) for (s, i) in res})
    if missing:
        cnt["NOT EXECUTED"] = cnt.get("NOT EXECUTED", 0) + len(missing)
        status = "incomplete"
        print("planned but no result (counted NOT EXECUTED):", ", ".join(f"{s}/{i}" for s, i in missing))
    ix = index()
    for e in ix:
        if e["run_id"] == rid:
            e.update(status=status, finished=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), counts=cnt, missing=[f"{s}/{i}" for s, i in missing])
    save_index(ix)
    write_summary(rid)
    print(rid, cnt)
    return cnt


def write_summary(run):
    """SUMMARY.md = the entry page of a run folder: what was run, the result, where everything is."""
    rid, d = resolve(run)
    e = next((x for x in index() if x["run_id"] == rid), {})
    plan = json.load(open(d / "plan.json", encoding="utf-8")) if (d / "plan.json").exists() else {}
    res = results_of(rid)
    cnt = e.get("counts") or {}
    rel = lambda p: p.relative_to(d).as_posix()
    L = [f"# Run {rid}", "",
         f"- Status: **{e.get('status', '?')}** · started {e.get('created', '')} · finished {e.get('finished', '')}",
         f"- Environment: {e.get('env') or plan.get('env', '')} · scope: `{' '.join(e.get('scope') or plan.get('scope', [])) or 'all'}` · catalog build: {plan.get('catalog_build', '')}",
         "- Result: " + (" · ".join(f"{k} {cnt[k]}" for k in ORDER if k in cnt) or "(none yet)") + f" (total {len(res)} ID"
         + (f", plus {len(plan['not_automated'])} NOT AUTOMATED not run" if plan.get("not_automated") else "") + ")", ""]
    bad = [(s, i, r) for (s, i), r in sorted(res.items()) if r.get("status") != "PASS"]
    if bad:
        L += ["## Cases not PASS", "", "| ID | Suite | Status | Cause hint | Note |", "|---|---|---|---|---|"]
        def why(r):                                            # case remark, else the first non-PASS step
            st = next((x for x in r.get("steps", []) if x.get("status") not in ("PASS", None)), {})
            return (r.get("remark") or st.get("remark") or st.get("actual") or "")[:120].replace("|", "/").replace("\n", " ")
        L += [f"| {i} | {s} | {r.get('status')} | {hint(r)} | {why(r)} |" for s, i, r in bad] + [""]
    L += ["## Where things are", "", "| What | Where |", "|---|---|"]
    for p in sorted((d / "excel").glob("*.xlsx")):
        if not p.name.startswith("~$"):
            L.append(f"| Excel deliverable | [{rel(p)}]({rel(p)}) |")
    for name, what in (("report.md", "Report A–F"), ("compare.md", "Comparison with earlier runs"), ("run_report.md", "Automatic report of autorun"),
                       ("gate.md", "Answers at the confirmation gate")):
        if (d / "reports" / name).exists():
            L.append(f"| {what} | [reports/{name}](reports/{name}) |")
    L += ["| Evidence images (original / per language) | `evidence/base/<suite>/<ID>/` · `evidence/<lang>/<suite>/<ID>/` |",
          "| Data for tools (plan, results, preflight) | `plan.json` · `data/` |",
          "| Cleanup / restore (registry, snapshots, role backup) | `safety/` |", "| Logs | `logs/` |", ""]
    (d / "SUMMARY.md").write_text("\n".join(L), encoding="utf-8")
    return d / "SUMMARY.md"


def consolidate(run):
    """Full evidence workbooks for a partial run. Per suite of the run's plan that has a spec: every case of the spec AS IT IS NOW; result =
    this run's, else the newest earlier *finished* run of the same suite (marked carried + src_run; spec_changed = the case hash of the catalog
    build that run was planned on differs from the current catalog, None = unknown), else none (NOT EXECUTED). Writes data/consolidated/
    results_<suite>.json, regenerates the suite's runtime translations (make_<lang>_runtime.py, if any) over the merged results and builds
    one workbook per language of qa.config.json language.excel. Run results are untouched."""
    import subprocess
    from . import catalog as cat_mod
    from ..reporting import spec as spec_mod
    rid, d = resolve(run)
    plan = json.load(open(d / "plan.json", encoding="utf-8")) if (d / "plan.json").exists() else {}
    ix = index()
    pos = next((k for k, e in enumerate(ix) if e["run_id"] == rid), len(ix))
    earlier = [e for e in ix[:pos] if e.get("status") == "finished"][::-1]          # newest first
    if plan.get("catalog_dir"):
        cat_mod.use(plan["catalog_dir"])                      # sandbox runs: compare against the sandbox catalog
    cat = cat_mod.load() or {"cases": {}}
    res_cache, hist_cache, out, notes = {}, {}, d / "data" / "consolidated", []

    def res(r):
        if r not in res_cache:
            try:
                res_cache[r] = results_of(r)
            except SystemExit:
                res_cache[r] = {}
        return res_cache[r]

    def spec_changed(build, tid):
        if build not in hist_cache:
            p = cat_mod.HISTORY / f"{build}.json"
            hist_cache[build] = json.load(open(p, encoding="utf-8"))["cases"] if p.exists() else None
        old, now = (hist_cache[build] or {}).get(tid), cat["cases"].get(tid)
        return None if not old or not now else old.get("hash") != now.get("hash")

    for it in plan.get("items", []):
        s, spec = it["suite"], it.get("spec")
        if not spec or not (WORK / spec).exists():
            continue
        ids = [str(k) for k in spec_mod.by_no(spec_mod.read(str(WORK / spec)))]
        merged, n_car, n_never = {}, 0, 0
        for i in ids:
            r = res(rid).get((s, i))
            if r:
                merged[i] = dict(r, src_run=rid, carried=False)
                continue
            for e in earlier:
                r = res(e["run_id"]).get((s, i)) if s in e.get("suites", []) else None
                if r:
                    pb = json.load(open(RUNS / e["run_id"] / "plan.json", encoding="utf-8")).get("catalog_build") if (RUNS / e["run_id"] / "plan.json").exists() else None
                    merged[i] = dict(r, src_run=e["run_id"], carried=True, spec_changed=spec_changed(pb, i) if pb else None)
                    n_car += 1
                    break
            else:
                n_never += 1
        if not n_car and not n_never:
            notes.append(f"{s}: the run covers every case of {Path(spec).name} -> its own workbook is already complete")
            continue
        out.mkdir(parents=True, exist_ok=True)
        (out / f"results_{s}.json").write_text(json.dumps(merged, ensure_ascii=False, indent=1), encoding="utf-8")
        env = dict(os.environ, QA_OUT=str(d), QA_RESULTS_DIR=str(out))
        sdir = WORK / "suites" / s
        for lang in cfg.langs:
            if (sdir / f"make_{lang}_runtime.py").exists():
                p = subprocess.run([sys.executable, "-m", f"suites.{s}.make_{lang}_runtime"], cwd=WORK, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
                if p.returncode:
                    notes.append(f"{s}: make_{lang}_runtime exit {p.returncode} (untranslated strings left): {(p.stdout + p.stderr).strip()[-300:]}")
        for lang, trf in ((lg, sdir / f"{lg}_text.py") for lg in cfg.langs):
            a = [sys.executable, "-m", "qakit.reporting.excel_builder", "--suite", s, "--spec", spec, "--lang", lang, "--consolidated"]
            if trf.exists():
                a += ["--tr", str(trf)]
            p = subprocess.run(a, cwd=WORK, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
            last = (p.stdout.strip().splitlines() or [""])[-1]
            notes.append(f"{s} {lang}: " + (last if p.returncode == 0 else f"exit {p.returncode}: {(p.stdout + p.stderr).strip()[-300:]}"))
        notes.append(f"{s}: {len(ids)} case = this run {len(ids) - n_car - n_never} · carried {n_car} · never run {n_never}")
    for n in notes:
        print(n)
    if pos < len(ix):
        write_summary(rid)
    return notes


def list_runs():
    print("| Run | Created | Scope | IDs | Status | Counts |\n|---|---|---|---|---|---|")
    for rid, d, mt in legacy():
        res = results_of(rid)
        cnt = {}
        for r in res.values():
            cnt[r.get("status", "?")] = cnt.get(r.get("status", "?"), 0) + 1
        print(f"| {rid} | {datetime.datetime.fromtimestamp(mt):%Y-%m-%d %H:%M} | (legacy, suites: {', '.join(sorted({s for s, _ in res}))}) | {len(res)} | legacy | {cnt} |")
    for e in index():
        print(f"| {e['run_id']} | {e['created']} | {' '.join(e.get('scope', []))} | {e['ids']} (+{e['not_automated']} not automated) | {e['status']} | {e.get('counts', '')} |")


def show(run):
    rid, d = resolve(run)
    res = results_of(rid)
    plan = json.load(open(d / "plan.json", encoding="utf-8")) if (d / "plan.json").exists() else {}
    print(f"# Run {rid} ({d})\n")
    print("| Suite | ID | Status | Cause hint | Remark |\n|---|---|---|---|---|")
    for (s, i), r in sorted(res.items(), key=lambda kv: (ORDER.index(kv[1].get("status")) if kv[1].get("status") in ORDER else 9, kv[0])):
        print(f"| {s} | {i} | {r.get('status')} | {hint(r)} | {(r.get('remark') or '')[:120].replace('|', '/')} |")
    if plan.get("not_automated"):
        print(f"\nNOT AUTOMATED (selected but no suite implements them, not run): {', '.join(x['id'] for x in plan['not_automated'])}")


def compare(run, base="previous", out=True):
    rid, d = resolve(run)
    cur = {i: (s, r) for (s, i), r in results_of(rid).items()}
    if base == "previous":                                     # per results file mtime: a legacy folder can mix old and new suites
        start = min((p.stat().st_mtime for p in result_files(d)), default=9e18)
        recs = []
        for b in [x[0] for x in legacy()] + [e["run_id"] for e in index()]:
            if b == rid:
                continue
            _, bd = resolve(b)
            for p in result_files(bd):
                if p.stat().st_mtime < start:
                    recs += [(p.stat().st_mtime, b, str(k), v) for k, v in json.load(open(p, encoding="utf-8")).items()]
        prev = {i: (b, r) for _, b, i, r in sorted(recs, key=lambda x: x[0])}   # oldest -> newest: later results overwrite
        base_label = "latest earlier result of every ID"
    else:
        bid, _ = resolve(base)
        prev = {i: (bid, r) for (s, i), r in results_of(bid).items()}
        base_label = bid
    cat = {"New FAIL (was PASS)": [], "Fixed (FAIL → PASS)": [], "Still FAIL": [], "Other status change": [], "First run": [], "Unchanged": []}
    for i, (s, r) in sorted(cur.items()):
        st = r.get("status")
        if i not in prev:
            cat["First run"].append(f"{i} {st}")
            continue
        b, pr = prev[i]
        ps = pr.get("status")
        line = f"{i}: {ps} ({b}) → {st}" + (f" · hint: {hint(r)}" if st != "PASS" else "")
        if ps == st:
            cat["Still FAIL" if st == "FAIL" else "Unchanged"].append(line if st == "FAIL" else f"{i} {st}")
        elif st == "FAIL" and ps == "PASS":
            cat["New FAIL (was PASS)"].append(line)
        elif st == "PASS" and ps == "FAIL":
            cat["Fixed (FAIL → PASS)"].append(line)
        else:
            cat["Other status change"].append(line)
    L = [f"# Comparison: {rid} ↔ {base_label}", ""]
    L += [f"- {k}: {len(v)}" for k, v in cat.items()] + [""]
    for k, v in cat.items():
        if v and k != "Unchanged":
            L += [f"## {k}", "", *[f"- {x}" for x in v], ""]
    txt = "\n".join(L)
    print(txt)
    if out and (d / "reports").is_dir() and not rid.startswith("legacy"):
        (d / "reports" / "compare.md").write_text(txt, encoding="utf-8")
    return cat


def main():
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest="cmd", required=True)
    n = sp.add_parser("new"); n.add_argument("--plan", required=True); n.add_argument("--label", default="")
    f = sp.add_parser("finish"); f.add_argument("run", nargs="?", default="latest"); f.add_argument("--status", default="finished", choices=["finished", "cancelled", "aborted"])
    sp.add_parser("list")
    sm = sp.add_parser("summary"); sm.add_argument("run", nargs="?", default="latest")
    s = sp.add_parser("show"); s.add_argument("run", nargs="?", default="latest")
    sp.add_parser("consolidate").add_argument("run", nargs="?", default="latest")
    c = sp.add_parser("compare"); c.add_argument("run", nargs="?", default="latest"); c.add_argument("--base", default="previous")
    a = ap.parse_args()
    if a.cmd == "new":
        new(a.plan, a.label)
    elif a.cmd == "finish":
        finish(a.run, a.status)
    elif a.cmd == "list":
        list_runs()
    elif a.cmd == "summary":
        print("written", write_summary(a.run))
    elif a.cmd == "show":
        show(a.run)
    elif a.cmd == "compare":
        compare(a.run, a.base)
    elif a.cmd == "consolidate":
        consolidate(a.run)


if __name__ == "__main__":
    main()
