"""The session state machine. Every command that does work asks `Session.require(...)` first, and a refusal (GateError) names the command that
comes next. This is where the process is enforced: hooks only stop shortcuts around it.

    ANALYZE -> CONNECT -> CONFIRM -> EXECUTE -> REPORT -> DONE        (INTAKE happens inside `session new`)

`waiting` is separate from the phase: while the assistant has asked the user something, nothing but `answer` / `status` / `ask` goes through.
"""
import difflib
import os
import re
import time
from contextlib import contextmanager

from . import SCHEMA_VERSION, analysis, events, memory, paths
from .store import SchemaError, locked, read_json, write_json

PHASES = ("ANALYZE", "CONNECT", "CONFIRM", "EXECUTE", "REPORT", "DONE")
MAX_REPAIRS = 3          # after the first run of a case, at most this many repair-and-rerun rounds
MIN_NOTE = 12            # characters a repair note's cause / evidence / change must have to say anything
SAME_CAUSE = 0.85        # a cause this close to an earlier one for the same case is the same cause
NEXT = {
    "ANALYZE": "oqa analyze --template > analysis.json, fill it in, then oqa analyze --submit analysis.json",
    "CONNECT": "oqa connect",
    "CONFIRM": "oqa strategy, then ask the user to approve it (the user runs oqa confirm, or presses Confirm in the dashboard)",
    "EXECUTE": "oqa run <case id>, in the order of strategy.md",
    "REPORT": "oqa report",
    "DONE": "nothing: the session is finished",
}


class GateError(Exception):
    """The command is not allowed in the session's current state. The message says why and what to run instead."""


class Session:
    def __init__(self, ws, session_id):
        self.ws, self.id = ws, session_id
        self.dir = paths.session_dir(ws, session_id)
        self.file = self.dir / "state.json"

    # --- opening and creating -------------------------------------------------
    @classmethod
    def open(cls, ws=None, session=None):
        ws, session = paths.resolve_session(ws, session)
        return cls(ws, session)

    @classmethod
    def create(cls, ws, testcase_file, cases_doc):
        """INTAKE is done: the workbook was read into `cases_doc`. Starts in ANALYZE and becomes the workspace's current session."""
        session = cls(ws, paths.new_session_id())
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        write_json(session.dir / "cases.json", cases_doc)
        write_json(session.file, {
            "schema_version": SCHEMA_VERSION, "id": session.id, "workspace": ws, "file": str(testcase_file), "created": now, "phase": "ANALYZE",
            "done": [{"phase": "INTAKE", "at": now}], "waiting": None, "strategy_sha": None, "confirmation": None, "attempts": {}, "runs": {}, "results": {}, "connect": None,
            "stopped": None, "hints": {}, "seen": 0, "timeouts": {}, "progress": {},
            "explore": "off" if os.environ.get("OQA_NO_EXPLORE_GATE") else "required",
        })
        (paths.workspace_dir(ws) / "current").write_text(session.id, encoding="utf-8")
        events.append(session.dir, "session_started", workspace=ws, file=str(testcase_file))
        events.append(session.dir, "phase", phase="ANALYZE", text=f"Read {len(cases_doc['cases'])} cases; analysing them")
        return session

    # --- state ------------------------------------------------------------------
    def state(self):
        data = read_json(self.file)
        if data is None:
            raise SchemaError(f"session {self.id} has no state.json", "session")
        return data

    @contextmanager
    def update(self):
        with locked(self.file):
            data = self.state()
            yield data
            write_json(self.file, data)

    def cases(self):
        return read_json(self.dir / "cases.json")

    # --- gates ------------------------------------------------------------------
    def require(self, *phases, working=True):
        """Raise GateError unless the session is in one of `phases` (and, for `working` commands, not waiting for the user)."""
        state = self.state()
        if working and state.get("stopped"):
            raise GateError("The user stopped the test. Do nothing until they resume it: run `oqa wait`.")
        if working and state["waiting"]:
            raise GateError(f"The session is waiting for the user to answer: \"{state['waiting']['question']}\". Do not continue until `oqa answer` is run with their reply.")
        if state["phase"] not in phases:
            raise GateError(f"Not allowed in phase {state['phase']} (allowed: {', '.join(phases)}). Next: {NEXT[state['phase']]}")
        return state

    def _enter(self, state, phase, text):
        state["phase"] = phase
        state["done"].append({"phase": phase, "at": time.strftime("%Y-%m-%d %H:%M:%S")})
        events.append(self.dir, "phase", phase=phase, text=text)

    # --- ANALYZE / CONNECT / CONFIRM ------------------------------------------------
    def analysis_submitted(self):
        """A valid analysis was saved. First time: go to CONNECT. A later correction (before confirming) drops the strategy so it is built again."""
        with self.update() as state:
            if state["phase"] not in ("ANALYZE", "CONNECT", "CONFIRM"):
                raise GateError(f"The analysis can no longer change in phase {state['phase']}.")
            if state["waiting"]:
                raise GateError("The session is waiting for the user to answer; run `oqa answer` first.")
            state["strategy_sha"] = None
            if state["phase"] == "ANALYZE":
                self._enter(state, "CONNECT", "Analysis accepted; testing the connection")

    def connected(self, info):
        """Called by `oqa connect` when every profile signed in. `info` (no secrets) is kept for the dashboard."""
        with self.update() as state:
            if state["phase"] != "CONNECT":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            state["connect"] = info
            self._enter(state, "CONFIRM", "Connection works; waiting for the user to approve the strategy")

    def strategy_ready(self, sha):
        with self.update() as state:
            if state["phase"] != "CONFIRM":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            state["strategy_sha"] = sha

    def confirm(self, by, sha):
        """The user's approval. `sha` must be the digest of the strategy on disk now, so what they approved is what runs."""
        with self.update() as state:
            if state["waiting"]:
                raise GateError("The session is waiting for the user to answer; run `oqa answer` first.")
            if state["phase"] != "CONFIRM":
                raise GateError(f"Nothing to confirm in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            if not state["strategy_sha"] or state["strategy_sha"] != sha:
                raise GateError("The strategy changed or was never built; run `oqa strategy` and ask the user to review it again.")
            state["confirmation"] = {"by": by, "at": time.strftime("%Y-%m-%d %H:%M:%S"), "strategy_sha": sha}
            events.append(self.dir, "confirmed", by=by, strategy_sha=sha)
            self._enter(state, "EXECUTE", "Confirmed; running the cases one by one")

    def confirm_current(self, by):
        """The user approves the strategy that is on disk now (the terminal and the dashboard both come here). Refuses a strategy that was
        edited or no longer matches the analysis."""
        strategy_doc = read_json(self.dir / "strategy.json")
        if not strategy_doc or not analysis.verify_strategy(strategy_doc, read_json(self.dir / "analysis.json")):
            raise GateError("strategy.json is missing or no longer matches the analysis; run `oqa strategy` again.")
        self.confirm(by, strategy_doc["sha"])
        return strategy_doc

    # --- asking the user ----------------------------------------------------------
    def ask(self, question, case=None):
        with self.update() as state:
            if state["phase"] == "DONE":
                raise GateError("The session is finished.")
            if state["waiting"]:
                raise GateError(f"Already waiting for the user: \"{state['waiting']['question']}\"")
            state["waiting"] = {"question": question, "case": case, "asked_at": time.strftime("%Y-%m-%d %H:%M:%S")}
            events.append(self.dir, "ask", question=question, **({"case": case} if case else {}))

    def answer(self, text):
        """The user replied. A question about one case gives that case a fresh repair budget and, if it was left PENDING / BLOCKED, lets it run again."""
        with self.update() as state:
            waiting = state["waiting"]
            if not waiting:
                raise GateError("Nobody is waiting for an answer.")
            case = waiting.get("case")
            state["waiting"] = None
            if case:
                state["attempts"].pop(case, None)
                state.setdefault("progress", {}).pop(case, None)
                if state["results"].get(case, {}).get("status") in ("PENDING", "BLOCKED"):
                    del state["results"][case]
                    if state["phase"] == "REPORT":
                        self._enter(state, "EXECUTE", f"{case} can run again after the answer")
            events.append(self.dir, "answered", answer=text, **({"case": case} if case else {}))
            return waiting

    # --- EXECUTE ------------------------------------------------------------------
    def begin_attempt(self, case):
        """Count one run of `case`. The first run plus MAX_REPAIRS reruns are allowed; the next is refused and the assistant must ask the user."""
        with self.update() as state:
            if state.get("stopped"):
                raise GateError("The user stopped the test. Do nothing until they resume it: run `oqa wait`.")
            if state["waiting"]:
                raise GateError(f"The session is waiting for the user: \"{state['waiting']['question']}\"")
            if state["phase"] != "EXECUTE":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            strategy = read_json(self.dir / "strategy.json") or {"order": []}
            if case not in strategy["order"]:
                raise GateError(f"{case} is not in the confirmed strategy.")
            if state["results"].get(case, {}).get("status") == "PASS":
                raise GateError(f"{case} already passed.")
            used = state["attempts"].get(case, 0)
            if used >= 1 + MAX_REPAIRS:
                raise GateError(f"{case} has used its first run and {MAX_REPAIRS} repairs. Stop fixing it: record it (oqa finish) or ask the user: oqa ask --case {case} \"...\"")
            if used >= 1:
                log = self._log(case)
                if log["last"] and log["last"]["status"] == "FAIL" and not log["noted"]:
                    raise GateError(f"{case} failed on run {log['last']['attempt']} and no repair note was recorded. Do not rerun on a guess: look at the picture and `oqa probe`, "
                                    f"then say what you found: oqa repair {case} --cause \"...\" --evidence \"...\" --change \"...\"")
            state["attempts"][case] = used + 1
            runs = state.setdefault("runs", {})                  # the run NUMBER never goes back, even when the user gives the case a fresh budget:
            runs[case] = runs.get(case, 0) + 1                   # two different runs must never share a number, a picture folder or a row in the report
            return runs[case]

    # --- repairing: what the assistant found, before it may run again -------------------------------
    def _log(self, case):
        """What the event log says about `case`'s runs: the last finished run, whether a repair note came after it, whether the page was probed since,
        every earlier run and every repair note."""
        log = {"last": None, "noted": False, "probed": False, "runs": [], "repairs": []}
        for e in events.read(self.dir):
            if e["kind"] in ("probe", "explore") and e.get("case") in (None, case):
                log["probed"] = True                 # the assistant looked at, or tried things on, the real page after the failure
            elif e.get("case") != case:
                continue
            elif e["kind"] == "attempt_finished":
                log.update(last=e, noted=False, probed=False)
                log["runs"].append(e)
            elif e["kind"] == "repair":
                log["noted"] = True
                log["repairs"].append(e)
        return log

    def awaiting_repair(self):
        """Cases whose last run failed and that have no repair note yet."""
        state = self.state()
        waiting = []
        for case in (read_json(self.dir / "strategy.json") or {"order": []})["order"]:
            log = self._log(case)
            if log["last"] and log["last"]["status"] == "FAIL" and not log["noted"] and case not in state["results"]:
                waiting.append(case)
        return waiting

    @staticmethod
    def _signature(text):
        return re.sub(r"\d+", "#", (text or "").lower()).strip()[:100]

    def repair(self, case, cause, evidence, change, reanalysis="", script_sha="", timeout_s=None, unchanged_reason=""):
        """The assistant's note after a failed run, required before the case runs again: the cause it found, how it knows, what it changed and, when the same failure
        came twice, how it now reads the case. Refuses a note that only repeats an earlier one or a repair that changed nothing."""
        cause, evidence, change, reanalysis, unchanged_reason = (x.strip() for x in (cause, evidence, change, reanalysis, unchanged_reason))
        with self.update() as state:
            if state.get("stopped"):
                raise GateError("The user stopped the test. Do nothing until they resume it: run `oqa wait`.")
            if state["waiting"]:
                raise GateError(f"The session is waiting for the user to answer: \"{state['waiting']['question']}\"")
            if state["phase"] != "EXECUTE":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            if case not in (read_json(self.dir / "strategy.json") or {"order": []})["order"]:
                raise GateError(f"{case} is not in the confirmed strategy.")
            log = self._log(case)
            last = log["last"]
            if not last or last["status"] != "FAIL" or case in state["results"]:
                raise GateError(f"{case} has no failed run waiting for a repair. Run it first: oqa run {case}")
            if log["noted"]:
                raise GateError(f"The repair note for run {last['attempt']} of {case} is already recorded. Next: oqa run {case}")
            for label, text in (("--cause", cause), ("--evidence", evidence), ("--change", change)):
                if len(text) < MIN_NOTE:
                    raise GateError(f"{label} says too little. Write what you actually found: the cause of the failure, what showed it (the picture, the probe, the page text) and what you changed.")
            plain = lambda text: re.sub(r"\W+", " ", text.lower()).strip()
            for earlier in log["repairs"]:
                if difflib.SequenceMatcher(None, plain(cause), plain(earlier["cause"])).ratio() >= SAME_CAUSE:
                    raise GateError(f"That is the cause you gave after run {earlier['attempt']}: \"{earlier['cause']}\". The change it led to did not fix the failure, so the cause is something else. "
                                    "Look again (the picture, `oqa probe`), or ask the user.")
            twice = len(log["runs"]) >= 2 and log["runs"][-2]["status"] == "FAIL" and self._signature(log["runs"][-2]["reason"]) == self._signature(last["reason"])
            if twice and len(reanalysis) < MIN_NOTE:
                raise GateError("The same failure came twice. Say how you now read the case (--reanalysis \"...\"): what the step really needs, what you assumed that was wrong.")
            if timeout_s is not None and not (analysis.MIN_TIMEOUT <= timeout_s <= analysis.MAX_TIMEOUT):
                raise GateError(f"--timeout is seconds from {analysis.MIN_TIMEOUT} to {analysis.MAX_TIMEOUT}.")
            if script_sha and last.get("script_sha") == script_sha and timeout_s is None and not unchanged_reason:
                raise GateError(f"The script of {case} is the same as the one that just failed. Change it, or say what else you changed (--timeout, or --unchanged \"data created through the UI\").")
            if timeout_s is not None:
                state.setdefault("timeouts", {})[case] = timeout_s
            note = {"case": case, "attempt": last["attempt"], "cause": cause, "evidence": evidence, "change": change, "probed": log["probed"]}
            for name, value in (("reanalysis", reanalysis), ("script_sha", script_sha), ("timeout_s", timeout_s), ("unchanged_reason", unchanged_reason)):
                if value not in ("", None):
                    note[name] = value
            events.append(self.dir, "repair", **note)
            memory.remember_repair(self.ws, memory.stem_of(state["file"]), case, note, self.id)
            return note

    def advance(self, case, step):
        """A run of `case` failed at `step`. When that is FURTHER than any earlier run got, it was not a failed repair: the case moved on and the next problem is a
        new one, so the run does not use up one of the repairs (the very first run of a case is always counted). Returns True when it got further."""
        with self.update() as state:
            best = state.setdefault("progress", {}).get(case, 0)
            if step <= best:
                return False
            state["progress"][case] = step
            if best and state["attempts"].get(case, 0) > 1:
                state["attempts"][case] -= 1
            return bool(best)

    def refund_attempt(self, case):
        """An attempt the user cut short does not count against the case's budget."""
        with self.update() as state:
            if state["attempts"].get(case):
                state["attempts"][case] -= 1
                state.setdefault("runs", {})[case] = max(0, state.get("runs", {}).get(case, 1) - 1)
                if not state["attempts"][case]:
                    del state["attempts"][case]

    # --- what only the user does: stop, resume, hint ------------------------------------
    def stop(self, by):
        """The user stops the test. The runner finishes the step it is on and does not start another; every working command is refused until `resume`."""
        with self.update() as state:
            if state["phase"] == "DONE":
                raise GateError("The session is finished.")
            if state.get("stopped"):
                raise GateError("Already stopped.")
            state["stopped"] = {"by": by, "at": time.strftime("%Y-%m-%d %H:%M:%S")}
            events.append(self.dir, "stopped", by=by)

    def resume(self, by):
        with self.update() as state:
            if not state.get("stopped"):
                raise GateError("The test is not stopped.")
            state["stopped"] = None
            events.append(self.dir, "resumed", by=by)

    def hint(self, case, text, by="dashboard"):
        """The user's advice for a case that failed (or was left BLOCKED / PENDING): it goes back in the queue with a fresh budget and without its old result,
        and the advice is shown to the assistant next time it runs the case. A question the assistant had asked about this case counts as answered."""
        text = (text or "").strip()
        if not text:
            raise GateError("Write the hint first.")
        with self.update() as state:
            if state["phase"] not in ("EXECUTE", "REPORT"):
                raise GateError(f"A hint is for a case that already ran; the session is in phase {state['phase']}.")
            if case not in (read_json(self.dir / "strategy.json") or {"order": []})["order"]:
                raise GateError(f"{case} is not in the confirmed strategy.")
            if state["results"].get(case, {}).get("status") == "PASS":
                raise GateError(f"{case} already passed.")
            state.setdefault("hints", {}).setdefault(case, []).append({"text": text, "at": time.strftime("%Y-%m-%d %H:%M:%S"), "by": by})
            state["attempts"].pop(case, None)
            state.setdefault("progress", {}).pop(case, None)
            state["results"].pop(case, None)
            if state["waiting"] and state["waiting"].get("case") == case:
                state["waiting"] = None
                events.append(self.dir, "answered", answer=text, case=case)
            events.append(self.dir, "hint", case=case, text=text)
            memory.remember_hint(self.ws, memory.stem_of(state["file"]), case, text, self.id)       # the next session on this file starts knowing it
            if state["phase"] == "REPORT":
                self._enter(state, "EXECUTE", f"{case} goes back in the queue after the user's hint")

    def news(self):
        """What the user did since the assistant was last told: confirmed, answered, hint, stopped, resumed (oldest first)."""
        seen = self.state().get("seen", 0)
        return [e for i, e in enumerate(events.read(self.dir)) if i >= seen and e["kind"] in events.USER_KINDS]

    def mark_seen(self):
        with self.update() as state:
            state["seen"] = len(events.read(self.dir))

    def finish_case(self, case, status, reason=""):
        """Record the end of a case. When every case of the strategy has a result, the session moves to REPORT."""
        if status not in events.CASE_STATUSES:
            raise GateError(f"status must be one of {', '.join(events.CASE_STATUSES)}")
        with self.update() as state:
            if state["phase"] != "EXECUTE":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            order = (read_json(self.dir / "strategy.json") or {"order": []})["order"]
            if case not in order:
                raise GateError(f"{case} is not in the confirmed strategy.")
            attempts = state["attempts"].get(case, 0)
            if not attempts and status in ("PASS", "FAIL"):
                raise GateError(f"{case} was never run; there is nothing to record.")
            state["results"][case] = {"status": status, "attempts": attempts, **({"reason": reason} if reason else {})}
            events.append(self.dir, "case_finished", case=case, status=status, attempts=attempts, **({"reason": reason} if reason else {}))
            memory.remember_result(self.ws, memory.stem_of(state["file"]), case, status, reason, self.id)
            if all(c in state["results"] for c in order):
                self._enter(state, "REPORT", "Every case has a result; building the report")

    def done(self, counts):
        with self.update() as state:
            if state["phase"] != "REPORT":
                raise GateError(f"Not allowed in phase {state['phase']}. Next: {NEXT[state['phase']]}")
            events.append(self.dir, "session_finished", counts=counts)
            self._enter(state, "DONE", "Report written; the session is finished")

    def declare(self, case, status, reason):
        """The assistant's own verdict when it stops working on a case: FAIL (the application is wrong: needs a run that failed with evidence),
        BLOCKED (cannot be done here) or PENDING (needs the user). PASS can only come from a run."""
        if status not in ("FAIL", "BLOCKED", "PENDING"):
            raise GateError("Only FAIL, BLOCKED or PENDING are declared; PASS comes from `oqa run`.")
        if not reason.strip():
            raise GateError("Say why (--reason).")
        if status == "FAIL":
            mine = [e for e in events.read(self.dir) if e.get("case") == case]
            attempts = [e for e in mine if e["kind"] == "attempt_finished"]
            if not attempts or attempts[-1]["status"] != "FAIL" or not any(e["kind"] == "step" and e["status"] == "FAIL" and e["images"] for e in mine):
                raise GateError(f"FAIL needs a run of {case} that failed with a picture of the failing step. Run it first.")
        self.finish_case(case, status, reason)

    def summary(self):
        state = self.state()
        order = (read_json(self.dir / "strategy.json") or {}).get("order", [])
        counts = {}
        for result in state["results"].values():
            counts[result["status"]] = counts.get(result["status"], 0) + 1
        return {"workspace": self.ws, "session": self.id, "file": state["file"], "phase": state["phase"], "waiting": state["waiting"], "stopped": state.get("stopped"),
                "confirmed": bool(state["confirmation"]), "cases_planned": len(order), "cases_done": len(state["results"]), "counts": counts,
                "attempts": state["attempts"], "next": NEXT[state["phase"]],
                "needs_repair": self.awaiting_repair() if state["phase"] == "EXECUTE" else []}
