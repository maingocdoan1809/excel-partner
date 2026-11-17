# -*- coding: utf-8 -*-
"""AUTORUN pipeline, started ONLY by the agent after the user's one-time confirmation at the gate of `/system-test run`
(skill reference execute.md). No scheduler, no other trigger.
  python -m qakit.tools.autorun --plan test_execution_output/runs/<run_id>/plan.json [--no-cleanup]
  python -m qakit.tools.autorun --suites <suite> ... [--no-cleanup]                       (ad-hoc, outputs in test_execution_output/adhoc/)
--plan: runs exactly the planned IDs per suite (matrix suites with --matrix [--targets]); re-executes itself with QA_OUT=<run folder>;
NOT AUTOMATED IDs are listed in the report; the run index is finished.
Pipeline: lock -> browser + login check (+ project data check) -> snapshot before -> every suite in its own process -> ALWAYS: cleanup,
snapshot after + diff -> Excel per suite (languages of qa.config.json) -> reports/run_report.md -> runs.finish + consolidated Excel.
Never asks anything. Whatever needs a human becomes PENDING (marked in the Excel and listed in the report with what is needed).
Exit code: 0 no FAIL, 1 any FAIL / integrity violation / script error, 2 nothing could run."""
import argparse, datetime, importlib, json, os, subprocess, sys, time
from ..core import browser, hooks, results, session
from ..core.config import WORK, cfg

sys.stdout.reconfigure(encoding="utf-8")
PY = sys.executable
LOCK = cfg.output_root / "_system" / "autorun.lock"     # one test browser -> one run at a time
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
LOG = cfg.logs_dir / f"autorun_{STAMP}.log"


def log(*a):
    line = f"[{datetime.datetime.now():%H:%M:%S}] " + " ".join(str(x) for x in a)
    print(line, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def sh(args, timeout=4 * 3600):
    r = subprocess.run([PY, "-m", *args], cwd=WORK, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)
    out = (r.stdout or "") + (r.stderr or "")
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"$ python -m {' '.join(args)}\n{out}\n")
    return r.returncode, out


def prepare_session():
    """None when ready, else a PENDING reason (needs a human)."""
    try:
        pg = browser.page()
        pg.goto(cfg.url(cfg.raw["app"].get("home_path", "/")), wait_until="domcontentloaded")
        pg.wait_for_timeout(1500)
        session.ensure_session(pg)
        if not session.logged_in(pg):
            return "not logged in after the login attempt: check auth settings / credentials"
        fn = getattr(hooks.module(), "check_data", None)
        if fn:
            msg = fn()
            log("data check:", msg)
            if msg and str(msg).startswith("ERROR"):
                return f"data check failed: {msg}"
        return None
    except session.NeedsHuman as e:
        return f"needs a human: {e}"
    except Exception as e:
        return f"cannot prepare the run: {str(e)[:250]}"
    finally:
        browser.stop()


def load_suite(name):
    if str(WORK) not in sys.path:
        sys.path.insert(0, str(WORK))
    return importlib.import_module(f"suites.{name}.cases").suite


def mark_all_pending(suites, reason, ids=None):
    for name in suites:
        s = load_suite(name)
        cfg.set_suite(name)
        results.path().parent.mkdir(parents=True, exist_ok=True)
        if results.path().exists():
            results.path().unlink()
        for n in (ids or {}).get(name) or sorted(s.cases, key=str):
            results.record(n, title="", status="PENDING", remark=reason, steps=[], started="", finished="", expected_steps=[])


def run_report(suites, pend_reason, integrity, notes, plan=None):
    cfg.reports_dir.mkdir(parents=True, exist_ok=True)
    path = cfg.reports_dir / ("run_report.md" if plan else f"run_report_{STAMP}.md")
    L = [f"# Automatic run report ({datetime.datetime.now():%Y-%m-%d %H:%M})", "", f"Project: {cfg.name} · {cfg.base_url} · log: `{LOG.name}`", ""]
    if plan:
        L += [f"Run: `{plan.get('run_id', '')}` · scope: `{' '.join(plan.get('scope', [])) or 'all'}` · catalog: {plan.get('catalog_build', '')}", ""]
    tot = {}
    sections = {"FAIL": [], "PENDING": [], "BLOCKED": [], "NOT EXECUTED": []}
    for name in suites:
        cfg.set_suite(name)
        for no, r in sorted(results.load().items(), key=lambda kv: str(kv[0])):
            st = r.get("status", "NOT EXECUTED")
            tot[st] = tot.get(st, 0) + 1
            if st in sections:
                needs = [x.get("needs") for x in r.get("steps", []) if x.get("needs")]
                sections[st].append(f"- `{name}` {no} {r.get('title', '')}: {r.get('remark', '')[:300]}" + (f" **NEEDS:** {needs[0]}" if needs else ""))
    L += ["## Totals", "", " · ".join(f"{k}: {v}" for k, v in sorted(tot.items())) or "(no case ran)", ""]
    if pend_reason:
        L += ["## ⛔ Could not run", "", pend_reason, ""]
    L += ["## Data integrity", "", integrity, ""]
    titles = {"PENDING": "What you need to do (PENDING)", "FAIL": "FAIL (application differs from the testcase)",
              "BLOCKED": "BLOCKED (cannot be observed / verified)", "NOT EXECUTED": "NOT EXECUTED (not automated yet or script error)"}
    for k in ("PENDING", "FAIL", "BLOCKED", "NOT EXECUTED"):
        if sections[k]:
            L += [f"## {titles[k]}", "", *sections[k], ""]
    if plan and plan.get("not_automated"):
        L += ["## NOT AUTOMATED (selected, no suite: not run)", "", *[f"- {x['id']} ({x['file']})" for x in plan["not_automated"]], ""]
    if notes:
        L += ["## Operational notes", "", *[f"- {n}" for n in notes], ""]
    path.write_text("\n".join(L), encoding="utf-8")
    return path, tot


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--suites", nargs="*", default=[])
    ap.add_argument("--no-cleanup", action="store_true")
    ap.add_argument("--plan", help="plan.json of tools.scope / tools.runs (selected IDs per suite); overrides --suites")
    a = ap.parse_args()
    plan, cmds = None, {}
    if a.plan:
        from .scope import command
        plan = json.load(open(a.plan, encoding="utf-8"))
        out = plan.get("out")
        if out and os.path.normcase(os.path.abspath(os.environ.get("QA_OUT", ""))) != os.path.normcase(os.path.abspath(out)):
            env = dict(os.environ, QA_OUT=out)                # cfg reads QA_OUT at import -> restart inside the run folder
            if plan.get("catalog_dir"):
                env["QA_SUITE_DIR"] = plan["catalog_dir"]
            return subprocess.run([PY, "-m", "qakit.tools.autorun", *sys.argv[1:]], cwd=WORK, env=env).returncode
        a.suites = [it["suite"] for it in plan["items"]]
        cmds = {it["suite"]: command(it)[2:] for it in plan["items"]}
    if not a.suites:
        ap.error("nothing to run: give --plan <plan.json> or --suites <suite> ...")
    cfg.ensure()
    LOCK.parent.mkdir(parents=True, exist_ok=True)
    if LOCK.exists() and time.time() - LOCK.stat().st_mtime < 6 * 3600:
        log("another autorun is active (lock younger than 6h):", LOCK)
        return 2
    LOCK.write_text(STAMP)
    notes, integrity, pend_reason, rc = [], "not checked", None, 0
    snap = cfg.raw["data"].get("snapshot", "none") != "none"
    try:
        log("suites:", a.suites, "| EXTERNAL allowed:", os.environ.get("QA_ALLOW_EXTERNAL") == "1")
        from ..core import events
        events.emit("phase", phase="run", text="running the selected cases")
        events.emit("run_started", run_id=(plan or {}).get("run_id", ""), suites=a.suites,
                    planned={it["suite"]: [str(i) for i in it["ids"]] for it in (plan or {}).get("items", [])})
        pend_reason = prepare_session()
        if pend_reason:
            log("PENDING ALL:", pend_reason)
            mark_all_pending(a.suites, pend_reason, {it["suite"]: it["ids"] for it in plan["items"]} if plan else None)
            integrity = "Nothing ran, so nothing changed."
            rc = 2
        else:
            if snap:
                sh(["qakit.reporting.snapshot", "take", "before"])
            for name in a.suites:
                cfg.set_suite(name)
                if results.path().exists():
                    results.path().unlink()
                code, out = sh(cmds.get(name) or [f"suites.{name}.cases"])
                log(f"suite {name}: exit {code}")
                if code != 0:
                    rc = 1
                    notes.append(f"suite {name} exited {code}: see log")
            events.emit("phase", phase="cleanup", text="removing the test data this run created")
            if not a.no_cleanup:
                code, out = sh(["qakit.tools.cleanup", "--apply"])
                log("cleanup:", out.strip().splitlines()[-1:] if out.strip() else code)
                if "MANUAL" in out or "FAILED" in out or code:
                    notes.append("cleanup NOT complete (MANUAL / FAILED items remain): see log")
            if snap:
                sh(["qakit.reporting.snapshot", "take", "after"])
                code, out = sh(["qakit.reporting.snapshot", "diff", "before", "after"])
                integrity = out.strip() or "no output"
                if code == 3:
                    notes.append("integrity UNVERIFIED: a snapshot is missing, so the data was NOT checked")
                elif code:
                    rc = 1
                    notes.append("integrity VIOLATION: objects that are not test objects changed (could also be real users working at the same time)")
            else:
                integrity = "data.snapshot = none in qa.config.json: the integrity of the data was NOT verified."
    except Exception as e:
        log("FATAL", repr(e))
        notes.append("runner error: " + repr(e)[:300])
        rc = 1
    finally:
        try:
            from ..core import events as _ev
            _ev.emit("phase", phase="report", text="building the Excel evidence and the run report")
        except Exception:
            pass
        try:
            for name in a.suites:
                s = load_suite(name)
                cfg.set_suite(name)
                for lang in cfg.langs:
                    args = ["qakit.reporting.excel_builder", "--suite", name, "--lang", lang]
                    if s.spec_path:
                        args += ["--spec", str(s.spec_path)]
                    trf = WORK / "suites" / name / f"{lang}_text.py"
                    if trf.exists():
                        args += ["--tr", str(trf)]
                    if plan:
                        args.append("--only-run")
                    code, out = sh(args)
                    notes.append(f"Excel {name} {lang}: " + (out.strip().splitlines()[-1][:200] if out.strip() else f"exit {code}"))
        except Exception as e:
            notes.append("cannot build Excel: " + repr(e)[:200])
        if plan and plan.get("run_id"):
            try:
                from . import runs
                notes += [f"consolidated Excel: {n}" for n in runs.consolidate(plan["run_id"])]
            except BaseException as e:
                notes.append("cannot build the consolidated Excel: " + repr(e)[:200])
        path, tot = run_report(a.suites, pend_reason, integrity, notes, plan)
        log("run report:", path, tot)
        try:
            from ..core import events
            events.emit("run_finished", run_id=(plan or {}).get("run_id", ""), counts=tot, rc=rc, notes=notes[-5:])
        except Exception:
            pass
        if plan and plan.get("run_id"):
            try:
                from . import runs
                runs.finish(plan["run_id"])
            except BaseException as e:
                log("run index not updated:", repr(e)[:200])
        LOCK.unlink(missing_ok=True)
    if tot.get("FAIL"):
        rc = rc or 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
