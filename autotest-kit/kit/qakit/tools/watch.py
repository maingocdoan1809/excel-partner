# -*- coding: utf-8 -*-
"""Stand guard over a run and COME BACK AT ONCE when something needs you, instead of waiting for the whole run to finish.
  python -m qakit.tools.watch [--run <run_id>] [--timeout 600]
Prints each new event as one line. Exit code: 10 = a case needs repair (run paused) or the run stopped itself - read the line, repair, trial that case;
0 = the run finished; 11 = nothing new for --timeout seconds (still running: call watch again). Default run = the newest one.
Typical use:  start `autorun` in the background, then `watch`: it returns on the FIRST problem, while the screen and the reasoning are fresh."""
import argparse, json, os, sys, time
from ..core.config import WORK, cfg

sys.stdout.reconfigure(encoding="utf-8")


def _file(run):
    root = WORK / cfg.raw["output"]["dir"]
    if run:
        for g in ("runs", "adhoc"):
            f = root / g / run.replace("adhoc-", "") / "events.jsonl"
            if f.exists():
                return f
        return None
    files = sorted((root / "runs").glob("*/events.jsonl"), key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


def line(j):
    k = j["kind"]
    t = j["ts"][11:]
    if k == "step":
        return f"{t} step {j.get('id')}#{j.get('n')} [{j.get('status')}] {str(j.get('caption', ''))[:60]} :: {str(j.get('actual', ''))[:140]}"
    if k == "case_finished":
        return f"{t} CASE {j.get('id')} => {j.get('status')} {str(j.get('remark', ''))[:160]}"
    if k == "phase":
        return f"{t} phase {j.get('phase')}: {str(j.get('text', ''))[:160]}"
    if k == "diagnosis":
        return f"{t} DIAGNOSIS {str(j.get('summary', ''))[:200]}  picture: {j.get('image', '')}"
    if k in ("run_paused", "run_stopped", "run_resumed"):
        return f"{t} {k.upper()}: {j.get('reason', '')}"
    return f"{t} {k}"


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--run")
    ap.add_argument("--timeout", type=int, default=600)
    a = ap.parse_args(argv)
    deadline = time.time() + a.timeout
    f, seen, state, first = None, 0, "running", True
    while time.time() < deadline:
        f = f or _file(a.run)
        if f and f.exists():
            lines = f.read_text(encoding="utf-8").splitlines()
            events = []
            for raw in lines[seen:]:
                try:
                    events.append(json.loads(raw))
                except ValueError:
                    continue
            for n, j in enumerate(events):
                k = j["kind"]
                # the state the log is in NOW: an old pause that was repaired and resumed must not stop a new watch
                if k == "run_paused":
                    state = "paused"
                elif k == "run_stopped":
                    state = "stopped"
                elif k == "run_finished":
                    state = "finished"
                elif k in ("run_started", "run_resumed", "case_started"):
                    state = "running"
                elif k == "phase" and str(j.get("text", "")).startswith("ASK THE USER"):
                    state = "ask"
                if k != "shot" and (not first or n >= len(events) - 6):          # the first look shows only the last few lines
                    print(line(j), flush=True)
            seen = len(lines)
            first = False
            if state in ("paused", "stopped", "ask"):
                return 10
            if state == "finished":
                return 0
        time.sleep(1.5)
    print("watch: nothing new for %d s; the run is still going" % a.timeout)
    return 11


if __name__ == "__main__":
    sys.exit(main())
