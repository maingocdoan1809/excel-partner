# -*- coding: utf-8 -*-
"""Tell everyone who watches (add-in, dashboard) what is happening NOW. Cheap: one line into the run's event stream.
  python -m qakit.tools.status <phase> "<what you are doing / found>" [--run <run_id>]
phase: setup | learn | automate | trial | preflight | plan | run | fix | cleanup | report | done | blocked
Goes to env QA_EVENTS / QA_OUT when set, else to --run, else to the newest run folder, else to today's ad-hoc folder."""
import argparse, sys
from ..core import events
from ..core.config import WORK, cfg

PHASES = ("setup", "learn", "automate", "trial", "preflight", "plan", "run", "fix", "cleanup", "report", "done", "blocked")


def _open_run_events(runs):
    """Events file of the newest run that has not ended yet (planned or running), else None."""
    import json
    if not runs.is_dir():
        return None
    for d in sorted((d for d in runs.iterdir() if (d / "events.jsonl").exists()), key=lambda d: (d / "events.jsonl").stat().st_mtime, reverse=True)[:1]:
        try:
            last = json.loads((d / "events.jsonl").read_text(encoding="utf-8").strip().splitlines()[-1])
        except Exception:
            return None
        return None if last.get("kind") in ("run_finished", "run_stopped") else d / "events.jsonl"
    return None


def target(run=None):
    import os
    if os.environ.get("QA_EVENTS") or os.environ.get("QA_OUT"):
        return events.path()
    runs = WORK / cfg.raw["output"]["dir"] / "runs"
    if run and (runs / run).is_dir():
        return runs / run / "events.jsonl"
    return _open_run_events(runs) or (cfg.out / "events.jsonl")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("phase", choices=PHASES)
    ap.add_argument("text")
    ap.add_argument("--run")
    a = ap.parse_args(argv)
    f = target(a.run)
    events.emit_to(f, "phase", phase=a.phase, text=a.text[:400])
    print(f"status {a.phase}: {a.text[:120]}  ->  {f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
