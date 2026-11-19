"""The ONE place that turns a session's files into "what is happening and what happened". The dashboard and the Excel report both read this, so
they can never disagree about a case, a step or a picture."""
import time

from . import events, knowledge, paths, recipe, refs
from .state import MAX_REPAIRS, PHASES, Session
from .store import read_json

CLOSED = ("PASS", "FAIL", "PENDING", "BLOCKED", "NOT EXECUTED")
IDLE_HOURS = 6              # a session that is not finished and has written nothing for this long was most likely left behind
LIVE = ("RUNNING", "REPAIRING", "EXPLORING")
STALE_S = 600               # nothing written for this long: a case shown as working is not (the assistant or the runner is gone), so it is not shown as working
TRY_S = 300                 # a `try` that started and never ended this long ago is not still running
WORK = ("try_started", "explore", "built", "case_started", "step", "attempt_finished")


def seconds_since(ts, now=None):
    try:
        return max(0, (now or time.time()) - time.mktime(time.strptime(ts, "%Y-%m-%d %H:%M:%S")))
    except (ValueError, TypeError):
        return 0


def quiet_seconds(state, log, now=None):
    """Seconds since the session last wrote anything (0 when finished)."""
    if state["phase"] == "DONE":
        return 0
    return seconds_since(log[-1]["ts"] if log else state["created"], now)


def idle_hours(state, log, now=None):
    """Hours since the session last did anything, or 0 when it is finished (a closed session is not idle, it is done)."""
    if state["phase"] == "DONE":
        return 0
    last = log[-1]["ts"] if log else state["created"]
    try:
        return max(0, ((now or time.time()) - time.mktime(time.strptime(last, "%Y-%m-%d %H:%M:%S"))) / 3600)
    except ValueError:
        return 0


def interrupted(records, hours, quiet_s=0):
    """In an idle session (or one that wrote nothing for STALE_S) a case shown as working is not: nobody is working on it. The report is not changed by this,
    only what the dashboard shows, so a spinner on the page always means something is really running."""
    if hours >= IDLE_HOURS or quiet_s >= STALE_S:
        for r in records:
            if r["status"] in LIVE:
                r["status"] = "INTERRUPTED"
            r.get("explore", {})["trying"] = None
            for a in r.get("attempts", []):
                if a["status"] == "RUNNING":
                    a["status"] = "INTERRUPTED"
                for s in a["steps"]:
                    if s["status"] == "RUNNING":
                        s["status"] = "PENDING"
    return records


def _steps(planned, ran, carried=None, running=False):
    """Every step the analysis planned, with what the attempt did to it; a step that never ran is PENDING (or SKIPPED when the attempt stopped earlier).
    A run that picked up at a later step did not run the steps before it: those show what an earlier run did, marked `carried`."""
    by_n = {e["n"]: e for e in ran}
    out, going = [], running
    for step in planned:
        e = by_n.get(step["n"])
        kept = None if e else (carried or {}).get(step["n"])
        e = e or kept
        now_running = going and not e                 # the attempt is on: the first step with no result yet is the one being run
        going = going and not now_running
        out.append({"n": step["n"], "caption": step["action"], "expected": step["expected"], "status": e["status"] if e else ("RUNNING" if now_running else "PENDING"),
                    "actual": e["actual"] if e else "", "images": e["images"] if e else [], "checks": e.get("checks") if e else None, "carried": bool(kept)})
    return out


def _explore(session, ident, mine, current, state, log):
    """Where the case is before it is run: which steps ran well on the real page (`verified`), the step being tried this very moment (`trying`, only between
    an `oqa try` starting and ending), whether the script is built and waits for a run (`built`), and whether the assistant is working on this case now (`live`)."""
    verified = sorted(recipe.steps_of(recipe.load(session, ident)))
    marks = [(i, e) for i, e in enumerate(mine) if e["kind"] in ("try_started", "explore", "built")]
    last_try = max((i for i, e in marks if e["kind"] in ("try_started", "explore")), default=-1)
    last_built = max((i for i, e in marks if e["kind"] == "built"), default=-1)
    trying = None
    if marks and marks[-1][1]["kind"] == "try_started" and seconds_since(marks[-1][1]["ts"]) < TRY_S and not state.get("stopped"):
        trying = {"step": marks[-1][1]["step"], "mode": marks[-1][1].get("mode", "")}
    built = last_built >= 0 and last_built > last_try          # an explore after the build means the case is being explored again
    started = any(e["kind"] == "case_started" for e in mine)
    live = bool(marks) and not built and not started and current == ident and not state.get("stopped") and state["phase"] == "EXECUTE" and seconds_since(log[-1]["ts"]) < STALE_S
    return {"verified": verified, "trying": trying, "built": built, "live": live}


def case_records(session: Session):
    """One record per case of the workbook, in the order the strategy runs them (cases held back come after)."""
    cases = session.cases()["cases"]
    analysis = {c["id"]: c for c in (read_json(session.dir / "analysis.json") or {"cases": []})["cases"]}
    strategy = read_json(session.dir / "strategy.json") or {"order": [], "excluded": []}
    state = session.state()
    log = events.read(session.dir)
    held = {e["id"]: e for e in strategy["excluded"]}
    by_id = {c["id"]: c for c in cases}
    ordered = [i for i in strategy["order"] if i in by_id]
    ordered += [i for i in analysis if i not in ordered and i in by_id] if analysis else []
    ordered += [c["id"] for c in cases if c["id"] not in ordered]
    out = []
    worked = [e for e in log if e["kind"] in WORK and e.get("case")]
    current = worked[-1]["case"] if worked else None          # the case the assistant touched last: the one being worked on, if any is
    for ident in ordered:
        case, item = by_id[ident], analysis.get(ident, {})
        mine = [e for e in log if e.get("case") == ident]
        attempts = []
        for n in sorted({e["attempt"] for e in mine if e["kind"] in ("case_started", "step", "attempt_finished")}):
            started = max((i for i, e in enumerate(mine) if e["kind"] == "case_started" and e["attempt"] == n), default=-1)
            # a run the user stopped and a later run of the same attempt number: only the end that comes after the latest start counts
            ended = next((e for i, e in reversed(list(enumerate(mine))) if e["kind"] == "attempt_finished" and e["attempt"] == n and i > started), None)
            from_step = mine[started].get("from_step", 1) if started >= 0 else 1
            carried = {}
            if from_step > 1:                         # the steps before it passed in an earlier run: the latest passing record of each
                for e in mine[:started]:
                    if e["kind"] == "step" and e["status"] == "PASS" and e["n"] < from_step:
                        carried[e["n"]] = e
            attempts.append({"n": n, "status": ended["status"] if ended else "RUNNING", "reason": ended["reason"] if ended else "", "failed_step": ended.get("failed_step") if ended else None,
                             "from_step": from_step,
                             "steps": _steps(item.get("steps") or [{"n": i, "action": s, "expected": ""} for i, s in enumerate(case["steps"], 1)],
                                             [e for e in mine if e["kind"] == "step" and e["attempt"] == n], carried, running=not ended and not state.get("stopped"))})
        explore = _explore(session, ident, mine, current, state, log)
        result = state["results"].get(ident)
        if result:
            status = result["status"]
        elif attempts:
            status = {"RUNNING": "RUNNING", "STOPPED": "STOPPED"}.get(attempts[-1]["status"], "REPAIRING")
        elif ident in strategy["order"]:
            status = "EXPLORING" if explore["live"] else ("BUILT" if explore["built"] else "PLANNED")
        elif ident in held:
            status = "HELD"
        else:
            status = "NEW"
        repairs = [{k: e[k] for k in ("ts", "attempt", "cause", "evidence", "change", "reanalysis", "probed", "timeout_s", "unchanged_reason") if k in e} for e in mine if e["kind"] == "repair"]
        out.append({"id": ident, "title": case["title"], "row": case["row"], "reuse": item.get("reuse", ""), "repairs": repairs, "timeout_s": state.get("timeouts", {}).get(ident) or item.get("timeout_s"),
                    "source": {"category": case.get("category", ""), "type": case.get("type", ""), "pre": case.get("pre", ""), "steps": case.get("steps", []),
                              "expected": case.get("expected", ""), "note": case.get("note", ""), "role": case.get("role", "")}, "profile": item.get("profile", ""), "verdict": item.get("verdict", held.get(ident, {}).get("verdict", "")),
                    "explore": explore, "status": status, "reason": (result or {}).get("reason") or held.get(ident, {}).get("reason", ""), "attempts": attempts,
                    "planned": [{"n": s["n"], "caption": s["action"], "expected": s["expected"]} for s in item.get("steps") or []],
                    "depends_on": item.get("depends_on") or [], "edge_cases": item.get("edge_cases") or [], "edge_none": item.get("edge_cases_none_reason", ""),
                    "data_needs": item.get("data_needs") or [], "clarity": (item.get("clarity") or {}).get("score"), "sufficiency": (item.get("sufficiency") or {}).get("score"),
                    "questions": item.get("questions") or [], "hints": (state.get("hints") or {}).get(ident, []), "used": state["attempts"].get(ident, 0),
                    "exhausted": state["attempts"].get(ident, 0) >= 1 + MAX_REPAIRS,
                    "can_hint": state["phase"] in ("EXECUTE", "REPORT") and ident in strategy["order"] and status != "PASS" and bool(attempts or result or state["attempts"].get(ident))})
    return out


def counts(records):
    out = {}
    for r in records:
        out[r["status"]] = out.get(r["status"], 0) + 1
    return out


def reports(session: Session):
    """The Excel evidence workbook written so far (one per session), or none."""
    folder = session.dir / "report"
    main = sorted(folder.glob("Evidence_*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True) if folder.is_dir() else []
    return [{"name": f.name, "size": f.stat().st_size, "main": True} for f in main[:1]]


def session_view(session: Session):
    state = session.state()
    log = events.read(session.dir)
    hours = idle_hours(state, log)
    records = interrupted(case_records(session), hours, quiet_seconds(state, log))
    current = state["phase"]
    reached = {d["phase"]: d["at"] for d in state["done"]}
    here = PHASES.index(current)
    phases = [{"phase": p, "at": reached.get(p, ""), "state": "current" if p == current else ("done" if i < here else "todo")} for i, p in enumerate(PHASES)]
    latest = next((e for e in reversed(log) if e["kind"] == "phase"), None)
    strategy_md = session.dir / "strategy.md"
    return {"workspace": session.ws, "session": session.id, "file": state["file"], "phase": current, "phases": phases, "waiting": state["waiting"], "stopped": state.get("stopped"), "can_stop": current != "DONE" and not state.get("stopped"),
            "confirmed": bool(state["confirmation"]), "created": state["created"], "updated": log[-1]["ts"] if log else state["created"], "idle_hours": round(hours, 1), "idle": hours >= IDLE_HOURS, "now": latest["text"] if latest else "",
            "counts": counts(records), "total": len(records), "cases": records, "connect": state.get("connect"),
            "reports": reports(session), "strategy": strategy_md.read_text(encoding="utf-8") if strategy_md.exists() else "", "can_confirm": current == "CONFIRM" and (session.dir / "strategy.json").exists()
            and state["strategy_sha"] is not None, "knowledge": len(knowledge.search(session.ws)),
            "refs": [{k: i.get(k) for k in ("name", "path", "kind", "by", "case")} for i in refs.items(session)]}


def list_sessions():
    """Every session of every workspace, newest first, with just enough to choose one."""
    out = []
    for ws in paths.workspace_names():
        for sid in paths.session_ids(ws):
            try:
                session = Session(ws, sid)
                state = session.state()
                hours = idle_hours(state, events.read(session.dir))
                records = interrupted(case_records(session), hours, quiet_seconds(state, events.read(session.dir)))
                out.append({"workspace": ws, "session": sid, "file": state["file"], "created": state["created"], "phase": state["phase"], "waiting": bool(state["waiting"]),
                            "idle": hours >= IDLE_HOURS, "idle_hours": round(hours, 1), "counts": counts(records), "total": len(records)})
            except Exception:
                continue                    # a session folder that is being written or damaged must not hide the others
    return sorted(out, key=lambda s: s["created"], reverse=True)
