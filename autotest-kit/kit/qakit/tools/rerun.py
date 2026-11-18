# -*- coding: utf-8 -*-
"""Run cases of a run AGAIN, inside THAT run: the new result REPLACES the old one. No new run, no new folder, no extra line in the run list.
  python -m qakit.tools.rerun <ID> [<ID> ...] [--run <run_id>]      the given cases (after you repaired them)
  python -m qakit.tools.rerun --pending [--run <run_id>]            the cases of the plan that have no result yet (what a pause left untouched)
  options: --no-cleanup (keep the test data this re-run creates)
Default run = the newest. Use this instead of `retry` while you are repairing a run: `retry` starts a NEW run, which costs a full set-up and leaves
a second run (and a dirtier history) behind. Start a new run only when the person asks for one.
What happens: the pictures of the case's earlier attempt are deleted, the case runs from its first step, its result record is replaced, the dashboard
shows the run as running again and the latest attempt as the case's state. The Excel evidence is NOT rebuilt here: run `result view=excel` once the run is
settled."""
import argparse, json, os, shutil, subprocess, sys
from ..core import events
from ..core.config import WORK, cfg
from . import runs, scope

PY = sys.executable


def _name(case):
    return f"TC{int(case):02d}" if str(case).isdigit() else str(case)


def _recorded(run_dir, suite):
    p = run_dir / "data" / f"results_{suite}.json"
    try:
        return set(json.load(open(p, encoding="utf-8")))
    except (OSError, ValueError):
        return set()


def select(plan, run_dir, wanted, pending):
    """[(item, ids)] of the plan: the named cases, or those with no result record yet."""
    chosen = []
    for item in plan.get("items", []):
        have = _recorded(run_dir, item["suite"])
        ids = [i for i in item["ids"] if (str(i) in wanted if wanted else (pending and str(i) not in have))]
        if ids:
            chosen.append((item, ids))
    return chosen


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("ids", nargs="*")
    ap.add_argument("--run", default="latest")
    ap.add_argument("--pending", action="store_true")
    ap.add_argument("--no-cleanup", action="store_true")
    a = ap.parse_args(argv)
    if not a.ids and not a.pending:
        ap.error("name the cases to run again, or --pending for the ones not reached yet")
    rid, run_dir = runs.resolve(a.run)
    plan_file = run_dir / "plan.json"
    if not plan_file.exists():
        sys.exit(f"{rid} has no plan.json: it is not a run made by `runs new`")
    plan = json.load(open(plan_file, encoding="utf-8"))
    if os.path.normcase(os.path.abspath(os.environ.get("QA_OUT", ""))) != os.path.normcase(os.path.abspath(run_dir)):
        env = dict(os.environ, QA_OUT=str(run_dir))               # cfg reads QA_OUT when it is imported: continue inside the run folder
        if plan.get("catalog_dir"):
            env["QA_SUITE_DIR"] = plan["catalog_dir"]
        return subprocess.run([PY, "-m", "qakit.tools.rerun", *(argv if argv is not None else sys.argv[1:]), "--run", rid], cwd=WORK, env=env).returncode
    chosen = select(plan, run_dir, {str(i) for i in a.ids}, a.pending)
    unknown = sorted({str(i) for i in a.ids} - {str(i) for _, ids in chosen for i in ids})
    if unknown:
        sys.exit(f"not in the plan of {rid}: {', '.join(unknown)}")
    if not chosen:
        print(f"{rid}: nothing to run again (every case of the plan has a result)")
        return 0
    todo = [str(i) for _, ids in chosen for i in ids]
    events_file = run_dir / "events.jsonl"
    events.emit_to(events_file, "run_resumed", suite="", reason=f"running again inside this run: {', '.join(todo)}", ids=todo)
    events.emit_to(events_file, "phase", suite="", phase="fix", text=f"running again: {', '.join(todo)} (the new result replaces the old one in this run)")
    rc = 0
    for item, ids in chosen:
        cfg.set_suite(item["suite"])
        for i in ids:                                                 # the earlier attempt's pictures must not outlive its result
            shutil.rmtree(cfg.evidence_dir / _name(i), ignore_errors=True)
        derived = dict(item, ids=list(ids))
        code = subprocess.run([PY, "-m", *scope.command(derived)[2:]], cwd=WORK, env=dict(os.environ, QA_OUT=str(run_dir))).returncode
        rc = rc or code
    if not a.no_cleanup:
        subprocess.run([PY, "-m", "qakit.tools.cleanup", "--apply"], cwd=WORK, env=dict(os.environ, QA_OUT=str(run_dir)))
    counts = runs.finish(rid)
    events.emit_to(events_file, "run_finished", suite="", run_id=rid, counts=counts, rc=rc)
    print(f"{rid}: {', '.join(todo)} run again in the same run -> {counts}")
    print("Excel evidence not rebuilt: run `/system-test result view=excel` when the run is settled.")
    return rc


if __name__ == "__main__":
    sys.exit(main())
