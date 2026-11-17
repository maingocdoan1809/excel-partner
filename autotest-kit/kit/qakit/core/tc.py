# -*- coding: utf-8 -*-
"""Testcase runtime: TC (steps, several images per step, status rules) + Suite runner.
  * A step without a real screenshot is never recorded as PASS (capture failure raises).
  * Status comes from assertions only: PASS / FAIL / BLOCKED / PENDING (needs a human) / NOT EXECUTED.
Suite module (suites/<name>/cases.py, cwd = autotest root):
    from qakit.core.tc import Suite
    suite = Suite("orders", "testcases/Orders_testcases.xlsx")
    @suite.case("ORD-001")
    def ord001(t, role, env):                 # matrix signature (role / env from the testcase role column), or def ord001(t) (plain)
        t.click("text=New order", "Click 「New order」")
        t.step(1, "Form opens", "Click New order", ok=..., actual="...", hl=["form"])
    if __name__ == "__main__":
        suite.main()                          # python -m suites.orders.cases [--matrix] [ID ...] [--targets=env:role,...]"""
import datetime, re, sys, traceback
from . import browser, diagnose, evidence, events, results, retry, session
from .retry import Diagnosed, Refused
from .breaker import Breaker
from .config import cfg
from .guard import GuardError

sys.stdout.reconfigure(encoding="utf-8")
M = cfg.raw["matrix"]


class Pending(Exception):
    """Raise inside a case when the rest of it REQUIRES a human (MFA, approval, physical action, external system).
    The case is recorded PENDING with `needs` = what the human must do."""
    def __init__(self, reason, needs=""):
        super().__init__(reason)
        self.reason, self.needs = reason, needs


REPAIR_STATUSES = ("BLOCKED", "NOT EXECUTED")


def needs_repair(steps):
    """True when the steps show a problem that is OURS to fix (data missing, screen refused, script error), not a verdict:
    BLOCKED / NOT EXECUTED steps, or a FAIL that was not declared an application defect with t.defect(). PENDING (a human is needed) is not repair work."""
    return any(s.get("status") in REPAIR_STATUSES or (s.get("status") == "FAIL" and not s.get("defect")) for s in steps)


def _problem_text(steps):
    s = next((x for x in steps if x.get("status") in REPAIR_STATUSES or (x.get("status") == "FAIL" and not x.get("defect"))), {})
    return f"{s.get('status', '')}: {s.get('caption', '')} :: {str(s.get('actual') or s.get('remark') or '')[:160]}".strip()


def _run_cfg(key, default):
    return cfg.raw.get("run", {}).get(key, default)


def _human(e):
    return isinstance(e, (Pending, session.NeedsHuman))


class TC:
    def __init__(self, no, title="", steps_text=None, spec=None):
        self.no, self.title, self.steps = no, title, []
        self.steps_text = steps_text or []
        self.spec = spec or {}
        self.started = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.pg = browser.page()
        self.pg.bring_to_front()
        self._pending = []
        self.dirty_role = False   # a case that changed the account's role itself: the runner re-applies the role before the next case
        self.timings = []         # per (env, role) run: {slot, start, end, seconds}
        self.ctx = ""             # '[label·role] ' prefix of every caption of the current (env, role) run
        events.emit("case_started", id=str(no), title=title, **self._where())

    def _where(self):
        sp = self.spec or {}
        return dict(sheet=sp.get("_sheet", ""), row=sp.get("_row", 0), header=sp.get("_header", 0), cols=sp.get("_cols", {}), file=str(sp.get("_file", "")))

    # ---- evidence ------------------------------------------------------------------------
    def snap(self, note, hl=None, cursor=None, wait=250, scroll=True, caption=None):
        """Extra image inside the step that is about to be recorded (numbered n-k)."""
        n, k = len(self.steps) + 1, len(self._pending) + 1
        try:
            ev = evidence.shot(self.no, n, self.ctx + (caption or note), hl=hl, pg=self.pg, wait=wait, scroll=scroll, k=k, note=self.ctx + note, cursor=cursor)
        except Exception as e:
            raise RuntimeError(f"evidence capture failed at {self.no} step {n}-{k} '{note}': {e}")
        self._pending.append(ev)
        events.emit("shot", id=str(self.no), n=len(self.steps) + 1, caption=note, image=ev["annotated"], **self._where())
        return ev

    def hover(self, selector, note=None):
        loc = evidence.as_locator(self.pg, selector).first
        loc.scroll_into_view_if_needed()
        loc.hover()
        self.pg.wait_for_timeout(300)
        return self.snap(note or f"Hover {selector}", hl=[selector], cursor=selector, scroll=False)

    def click(self, selector, note=None, image_before=True):
        """Image BEFORE the click (element boxed + pointer), then click."""
        loc = evidence.as_locator(self.pg, selector).first
        loc.scroll_into_view_if_needed()
        if image_before:
            loc.hover()
            self.pg.wait_for_timeout(200)
            self.snap(note or f"Click {selector}", hl=[selector], cursor=selector, scroll=False)
        loc.click()
        self.pg.wait_for_timeout(500)

    def fill(self, selector, value, note=None):
        """Type into a field, image AFTER typing (field boxed)."""
        loc = evidence.as_locator(self.pg, selector).first
        loc.scroll_into_view_if_needed()
        loc.fill(value)
        shown = value if len(str(value)) <= 30 else str(value)[:27] + "…"
        return self.snap(note or f"Enter 「{shown}」", hl=[selector], scroll=False)

    # ---- repair: the testcase steps are the IDEAL path; these helpers deal with what gets in the way --------------------------------------------
    def submit(self, selector, note=None, ok=None, fix=None):
        """Click a save / submit control and PROVE it worked. If the screen refuses (error banner, a field flagged invalid, a required field that was
        empty before and still is) the kit takes a picture of that moment, reads the messages, fills the EMPTY required fields with valid test values
        (or calls `fix(diagnosis)`), and clicks again. Raises Diagnosed (with the screen's own words) when it is still refused; that is a problem
        of the TEST (input / data), never a verdict on the application. `ok()` = extra proof that the save took effect (e.g. the record is listed)."""
        def go():
            before = {f["label"] for f in diagnose.read(self.pg)["fields"] if f["why"] == "required and empty"}
            self.click(selector, note or f"Click {selector}")
            self.pg.wait_for_timeout(700)
            d = diagnose.read(self.pg)
            flagged = [f for f in d["fields"] if f["why"] != "required and empty"]
            still_empty = [f for f in d["fields"] if f["why"] == "required and empty" and f["label"] in before]
            if d["errors"] or flagged or still_empty:
                raise Refused(d["summary"])
            if ok is not None and not ok():
                raise Refused("the save was accepted but its expected result is not on the screen")
        retry.attempt(f"save ({note or selector})", go, fix=fix or (lambda d: diagnose.autofill(self.pg, d)), pg=self.pg, case=str(self.no))

    def ensure(self, n, label, present, create, fix=None):
        """The testcase needs something on the screen (a record, a row, a tag) - MAKE it when it is not there, as test data (names start with the test prefix).
        present() -> bool; create() makes it through the UI. Missing and creatable -> created, checked again, True. Not creatable after the attempts ->
        a BLOCKED step carrying the real reason, False (the case returns). Never mark a case BLOCKED because data was missing without trying this."""
        if present():
            return True
        events.emit("phase", phase="fix", text=f"{self.no}: '{label}' is not on the screen; creating it as test data")

        def make():
            create()
            if not present():
                raise Refused(f"'{label}' is still not on the screen after creating it")
        try:
            retry.attempt(f"create: {label}", make, fix=fix or (lambda d: diagnose.autofill(self.pg, d)), pg=self.pg, case=str(self.no))
            return True
        except (Pending, session.NeedsHuman):
            raise
        except Exception as e:
            self.blocked(n, f"Precondition: {label}", "create the missing test data", f"could not create it: {str(e)[:300]}")
            return False

    def need_input(self, n, caption, action, what):
        """A step cannot be checked because INPUT is missing (a design document, a value the testcase does not give). Records PENDING with what is
        needed AND raises the question where the assistant is watching: `watch` returns at once, and the assistant must ASK THE USER (one short question
        naming the step and the document) before it reports anything. PENDING without having asked is not an acceptable ending."""
        self.pending(n, caption, action, f"input missing: {what}", needs=what)
        events.emit("phase", phase="blocked", text=f"ASK THE USER: {self.no} step {n} needs: {what}"[:400])
        print(f"!! ASK THE USER: case {self.no} step {n} needs: {what}")

    def reach(self, n, label, condition, grow, rounds=5, fix=None):
        """The testcase needs a CONDITION the data on screen does not show yet (tags overflowing their area, a list longer than one page, a value at its
        maximum length). Do not stop at PENDING / BLOCKED: grow the data round by round until the condition holds. condition() -> bool; grow(i) adds one
        more step toward it (i = 0, 1, 2 ... e.g. one more tag, one more record). True when reached. After `rounds` rounds it records BLOCKED with how far it
        got and returns False."""
        if condition():
            return True
        for i in range(max(1, int(rounds))):
            events.emit("phase", phase="fix", text=f"{self.no}: '{label}' not reached yet; adding more test data (round {i + 1}/{rounds})")
            try:
                retry.attempt(f"reach: {label} (round {i + 1})", lambda: grow(i), fix=fix or (lambda d: diagnose.autofill(self.pg, d)), pg=self.pg, case=str(self.no))
            except (Pending, session.NeedsHuman):
                raise
            except Exception as e:
                self.blocked(n, f"Condition: {label}", "grow the test data", f"could not add more data in round {i + 1}: {str(e)[:260]}")
                return False
            if condition():
                return True
        self.blocked(n, f"Condition: {label}", "grow the test data", f"still not reached after {rounds} rounds of adding data")
        return False

    def defect(self, n, caption, action, expected, actual, ruled_out, hl=None):
        """The ONLY way to record FAIL for 'the application is wrong'. `ruled_out` lists what you checked to show the fault is NOT in the test:
        e.g. ["inputs are valid per the rule", "the data exists (3 rows listed)", "no validation message on the screen", "same result after a reload"].
        A picture of the screen and its messages are saved with the step. Without ruled_out it raises ValueError: a failure you cannot defend is not a FAIL."""
        if not ruled_out:
            raise ValueError("defect(): ruled_out must list what you checked to show the fault is the application's, not the test's (input, data, selector, environment)")
        d = diagnose.snapshot(self.pg, f"{self.no}-defect", case=str(self.no))
        self.step(n, caption, action, ok=False, actual=f"expected: {expected} | actual: {actual}", hl=hl,
                  remark="APPLICATION DEFECT. Ruled out: " + "; ".join(ruled_out) + f" | screen: {d['summary']}")
        self.steps[-1]["defect"] = True

    def step(self, n, caption, action, ok, actual, hl=None, remark="", status=None, wait=400, scroll=True):
        """FINAL image of the step (after any snap / hover / click images) + verdict."""
        idx, k = len(self.steps) + 1, (len(self._pending) + 1 if self._pending else None)
        if not [x for x in (hl or []) if x]:
            # a picture without a red box proves nothing: say so where the assistant and the person both look (log, dashboard)
            remark = (remark + " | " if remark else "") + "NO RED BOX: pass hl=[the element this step is about]"
            print(f"   WARNING step {idx} of {self.no} has no red box: pass hl=[selector of the element the step is about]")
        try:
            ev = evidence.shot(self.no, idx, self.ctx + caption, hl=hl, pg=self.pg, wait=wait, scroll=scroll, k=k, note=self.ctx + caption)
        except Exception as e:
            raise RuntimeError(f"evidence capture failed at {self.no} step '{caption}': {e}")
        images, self._pending = self._pending + [ev], []
        st = status or ("PASS" if ok else "FAIL")
        self.steps.append(dict(orig_step=n, caption=self.ctx + caption, action=action, actual=actual, status=st, remark=remark, images=images,
                               raw=ev["raw"], annotated=ev["annotated"], url=ev["url"], ts=ev["ts"]))
        print(f"  {self.no} step{len(self.steps)} [{st}] {caption} :: {actual}" + (f"  ({len(images)} images)" if len(images) > 1 else ""))
        events.emit("step", id=str(self.no), n=len(self.steps), total=len(self.steps_text), status=st, caption=caption, actual=str(actual)[:300],
                    remark=str(remark)[:400], images=[i["annotated"] for i in images], **self._where())

    def _nonimage(self, n, caption, action, actual, status, remark, **kw):
        self.steps.append(dict(orig_step=n, caption=self.ctx + caption, action=action, actual=actual, status=status, remark=remark,
                               raw=None, annotated=None, url="", ts="", images=[], **kw))
        print(f"  {self.no} step{len(self.steps)} [{status}] {caption} :: {actual}")
        events.emit("step", id=str(self.no), n=len(self.steps), total=len(self.steps_text), status=status, caption=caption, actual=str(actual)[:300],
                    images=[], **self._where())

    def blocked(self, n, caption, action, reason):
        """Step that cannot be verified (environment / data / observability limit). State the condition that would unblock it."""
        self._nonimage(n, caption, action, reason, "BLOCKED", reason)

    def todo(self, n, caption, action=""):
        """Placeholder of a step whose script is not written yet (generated skeletons): NOT EXECUTED, never PASS."""
        self._nonimage(n, caption, action, "script not written yet (generated skeleton)", "NOT EXECUTED", "TODO")

    def pending(self, n, caption, action, reason, needs=""):
        """Step that needs a human (PENDING, never PASS). `raise Pending(...)` stops the whole case."""
        self._nonimage(n, caption, action, reason, "PENDING", reason + (f" | NEEDS: {needs}" if needs else ""), needs=needs)

    def verdict(self, remark="", force=None, covered=()):
        """(status, remark) of the case from its recorded steps. covered = expected step numbers a recorded step stands for
        (one step() may cover several lines of the testcase)."""
        sts = [s["status"] for s in self.steps]
        overall = force or ("FAIL" if "FAIL" in sts else "PENDING" if "PENDING" in sts else "BLOCKED" if "BLOCKED" in sts
                            else "NOT EXECUTED" if ("NOT EXECUTED" in sts or not sts) else "PASS")
        if overall == "PASS" and not force and self.steps_text:
            done = {s.get("orig_step") for s in self.steps} | set(covered)
            missing = [i for i in range(1, len(self.steps_text) + 1) if i not in done]
            if missing:
                overall = "NOT EXECUTED"
                remark = (remark + " | " if remark else "") + f"PASS withheld: expected step(s) {missing} were not recorded"
        if not remark and overall in ("FAIL", "BLOCKED", "PENDING"):
            bad = next((s for s in self.steps if s["status"] == overall and (s.get("remark") or s.get("actual"))), None)
            remark = str((bad or {}).get("remark") or (bad or {}).get("actual") or "")[:300]
        return overall, remark

    def finish(self, remark="", force=None, covered=()):
        overall, remark = self.verdict(remark, force, covered)
        events.emit("case_finished", id=str(self.no), status=overall, remark=remark[:300], steps=len(self.steps), **self._where())
        results.record(self.no, title=self.title, status=overall, remark=remark, steps=self.steps, started=self.started,
                       finished=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), expected_steps=self.steps_text, timings=self.timings)
        print(f"== {self.no} => {overall}")
        return overall


def _kw_match(text, words):
    t = str(text or "")
    for w in words:
        if re.search(r"(?<![\w])" + re.escape(w) + r"(?![\w])", t, flags=re.I if w.isascii() else 0):
            return True
    return False


def recognised(role_text):
    """(env keys, role keys) whose keywords appear in the role / environment column (empty lists = nothing recognised)."""
    return ([k for k, ws in M.get("envs", {}).items() if _kw_match(role_text, ws)], [k for k, ws in M.get("roles", {}).items() if _kw_match(role_text, ws)])


def targets_of(role_text):
    """(env, role) slots a case must run in, from its role / environment column. Keywords: matrix.envs / matrix.roles of the config.
    Nothing recognised -> (default_env, default_role)."""
    envs, roles = recognised(role_text)
    return [(e, r) for e in (envs or [M.get("default_env", "default")]) for r in (roles or [M.get("default_role", "default")])]


def _label(env, role):
    """Caption prefix of a matrix run: [env·role], [role] when the project has one environment, '' when it has neither."""
    if M.get("envs"):
        return f"[{session.ENV_LABEL.get(env, env)}·{role}] "
    return f"[{role}] " if M.get("roles") else ""


class Suite:
    """A set of testcases bound to one spec workbook. Registers case functions and runs them by ID."""
    def __init__(self, name, spec_path=None):
        from ..reporting import spec as spec_mod
        self.name = name
        self.spec_path = spec_path
        self.roles = {}          # plain suites: case ID -> role to run it with
        self.setup_fn = None     # fn(tcs, suite): prepare test data before the matrix loop
        self.teardown_fn = None  # fn(suite): remove that data after roles / environment were restored
        self.state = {}
        self.breaker = Breaker()
        cfg.set_suite(name)
        from .config import WORK
        sp = None
        if spec_path:
            from pathlib import Path
            sp = Path(spec_path) if Path(spec_path).is_absolute() else WORK / spec_path
        self.spec = spec_mod.by_no(spec_mod.read(sp)) if sp else {}
        self.cases = {}
        for c in self.spec.values():
            c["_file"] = str(sp) if sp else ""

    targets_of = staticmethod(targets_of)

    def _stop_reason(self, t, err=None):
        """Feed one finished case to the circuit breaker; the reason text when the run must stop. `err` = the exception text of a
        case that crashed (its recorded steps may all be PASS, but the case did not finish)."""
        if err:
            return self.breaker.observe("NOT EXECUTED", [], err)
        st, rem = t.verdict()
        return self.breaker.observe(st, t.steps, rem)

    def _diag(self, t, e, n):
        """What the screen says about a case that crashed: validation messages, invalid fields, a picture (diagnose/<time>_<ID>-error.png)."""
        try:
            d = e.diagnosis if isinstance(e, Diagnosed) else diagnose.snapshot(t.pg if t else browser.page(), f"{n}-error", error=e, case=str(n))
            return d.get("summary", "")
        except Exception:
            return ""

    def _pause_reason(self, t):
        """run.stop_on_problem: the moment a case needs repair, stop the run so it is repaired NOW, while the screen and the reasoning are fresh,
        instead of finishing a whole suite on broken ground and repeating it."""
        if _run_cfg("stop_on_problem", False) and needs_repair(t.steps):
            return f"case {t.no} needs repair: {_problem_text(t.steps)}"
        return None

    def _pause(self, reason, remaining):
        """Leave the cases not reached untouched (PLANNED); `retry` / a trial of that case picks up after the repair."""
        print(f"!! RUN PAUSED FOR REPAIR: {reason}")
        events.emit("run_paused", reason=reason[:400], remaining=[str(n) for n in remaining])

    def _call_case(self, t, n, role, env):
        """One execution of a case function with the runner's error handling (used for the first try and the immediate retry)."""
        try:
            self.cases[n](t, role, env) if self._matrix_fn(n) else self.cases[n](t)
        except (Pending, session.NeedsHuman) as p:
            t.pending(len(t.steps) + 1, "Needs a human", "pending", getattr(p, "reason", str(p)), needs=getattr(p, "needs", ""))
        except GuardError as e:
            t.pending(len(t.steps) + 1, "Guard", "guard", f"needs the user's permission: {str(e)[:220]}", needs="allow the action refused by the guard")
        except Exception as e:
            traceback.print_exc()
            screen = self._diag(t, e, n)
            t._nonimage(len(t.steps) + 1, "runner error", "", "script / runner error (not a verdict on the application): " + str(e)[:200]
                        + (f" | screen says: {screen}" if screen else ""), "NOT EXECUTED", str(e)[:200])

    def _run_case_settled(self, t, n, role, env):
        """Run the case; when it ends needing repair, run it once more AT ONCE (run.case_retries, default 1) from a clean page before going on."""
        before = len(t.steps)
        retries = max(0, int(_run_cfg("case_retries", 0)))
        for attempt_no in range(retries + 1):
            self._call_case(t, n, role, env)
            if attempt_no == retries or not needs_repair(t.steps[before:]):
                return
            events.emit("phase", phase="fix", text=f"{n}: {_problem_text(t.steps[before:])}; running this case once more before going on")
            del t.steps[before:]
            t._pending = []
            try:
                t.pg = browser.page(); t.pg.reload()
            except Exception:
                pass
            events.emit("case_started", id=str(n), title=t.title, **t._where())

    def _stop(self, reason, remaining):
        """The breaker tripped: every case not run yet is recorded NOT EXECUTED with the reason, then the run ends."""
        print(f"!! RUN STOPPED: {reason}")
        events.emit("run_stopped", reason=reason, remaining=[str(n) for n in remaining])
        for n in remaining:
            sp = self.spec.get(n, {})
            events.emit("case_finished", id=str(n), status="NOT EXECUTED", remark=f"run stopped, not run: {reason}"[:300], steps=0,
                        sheet=sp.get("_sheet", ""), row=sp.get("_row", 0), header=sp.get("_header", 0), cols=sp.get("_cols", {}), file=str(sp.get("_file", "")))
            self._record_partial(n, None, "NOT EXECUTED", f"run stopped, not run: {reason}")

    def case(self, no, role=None):
        def deco(fn):
            self.cases[no] = fn
            if role:
                self.roles[no] = role
            return fn
        return deco

    def new_tc(self, no):
        sp = self.spec.get(no, {})
        return TC(no, str(sp.get("view", "")), sp.get("act_lines", []), sp)

    def _session_or_pending(self, nos):
        """Make sure the test browser is logged in; when only a human can do it every case becomes PENDING."""
        try:
            session.ensure_session(browser.page())
            return True
        except Exception as e:
            msg = str(e)[:250]
            for n in nos:
                self._record_partial(n, None, "PENDING" if _human(e) else "NOT EXECUTED", ("NEEDS: " if _human(e) else "login failed: ") + msg)
            return False

    def run(self, nos):
        nos = sorted(nos, key=lambda n: (self.roles.get(n) or "", str(n)))
        cur = None
        try:
            if not self._session_or_pending(nos):
                return
            for n in nos:
                t = None
                err = None
                try:
                    t = self.new_tc(n)
                    role = self.roles.get(n)
                    if role and role != cur:
                        session.switch_role(t.pg, role)
                        cur = role
                    self.cases[n](t)
                except (Pending, session.NeedsHuman) as p:
                    reason = getattr(p, "reason", str(p))
                    self._record_partial(n, t, "PENDING", f"{reason} | NEEDS: {p.needs}" if getattr(p, "needs", "") else reason)
                except GuardError as e:
                    self._record_partial(n, t, "PENDING", f"needs the user's permission: {str(e)[:250]}")
                except Exception as e:
                    traceback.print_exc()
                    err = f"{type(e).__name__}: {str(e)[:200]}"
                    screen = self._diag(t, e, n)
                    self._record_partial(n, t, "NOT EXECUTED", "script / runner error (not a verdict on the application): " + traceback.format_exc()[-300:]
                                         + (f" | screen says: {screen}" if screen else ""))
                if t is not None and (reason := self._stop_reason(t, err)):
                    self._stop(reason, nos[nos.index(n) + 1:])
                    break
                if t is not None and (paused := self._pause_reason(t)):
                    self._pause(paused, nos[nos.index(n) + 1:])
                    break
        finally:
            if cur:
                try:
                    session.switch_role(browser.page(), "home")
                except Exception:
                    traceback.print_exc()
                    print("!! default role NOT restored:", session.role_command("home"))
            browser.stop()

    # Matrix runner: every case runs for EVERY (env, role) of its role column, batched by (env, role) so switches happen as rarely as
    # possible. All runs of one case accumulate in ONE testcase record. Case signature: fn(t, role, env).
    def run_matrix(self, nos=None, only_targets=None, restore=True):
        nos = list(nos) if nos else sorted(self.cases)
        plan = {}
        for n in nos:
            for tg in targets_of(self.spec.get(n, {}).get("role")):
                if only_targets is None or tg in only_targets:
                    plan.setdefault(tg, []).append(n)
        env_order = list(M.get("envs", {})) + [M.get("default_env", "default")]
        role_order = list(M.get("roles", {})) + [M.get("default_role", "default")]
        order = sorted(plan, key=lambda tg: (env_order.index(tg[0]) if tg[0] in env_order else 99, role_order.index(tg[1]) if tg[1] in role_order else 99))
        tcs = self.tcs = {}
        if not self._session_or_pending(nos):
            browser.stop()
            return
        pg0 = browser.page()
        orig = session.role_env_state()
        bk = session.backup_roles()
        print("original state:", orig, "| backup:", bk)
        self.selected = list(nos)              # setup_fn can prepare only the data the selected cases need
        try:
            if self.setup_fn:
                self.setup_fn(tcs, self)
            stopped = None
            for env, role in order:
                if stopped:
                    break
                try:
                    session.switch_env(env, roles=[r for e, r in plan if e == env])
                    session.switch_role(pg0, role)
                except Exception as e:
                    print("!! cannot prepare", env, role, e)
                    for n in plan[(env, role)]:
                        t = tcs.setdefault(n, self.new_tc(n)); t.ctx = _label(env, role)
                        t.pending(len(t.steps) + 1, "Prepare environment / role", "switch", f"cannot switch environment / role: {str(e)[:200]}",
                                  needs="configure roles.switch / envs.switch, or switch by hand")
                    continue
                for n in plan[(env, role)]:
                    if stopped:
                        break
                    t = tcs.setdefault(n, self.new_tc(n))
                    t.ctx = _label(env, role)
                    t.pg = browser.page(); t.pg.bring_to_front()
                    t0 = datetime.datetime.now()
                    try:
                        self._run_case_settled(t, n, role, env)
                    finally:
                        t1 = datetime.datetime.now()
                        t.timings.append(dict(slot=_label(env, role).strip(" []") or "default", start=t0.strftime("%H:%M:%S"), end=t1.strftime("%H:%M:%S"),
                                              seconds=round((t1 - t0).total_seconds(), 1)))
                        if t.dirty_role:
                            session._state["role"], t.dirty_role = None, False
                        results.record(n, title=t.title, status="NOT EXECUTED", remark="(running)", steps=t.steps, started=t.started,
                                       finished=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), expected_steps=t.steps_text)
                    stopped = self._stop_reason(t)
                    if stopped:
                        done = set(tcs)
                        self._stop(stopped, [x for tg in order for x in plan[tg] if x not in done])
                    elif (paused := self._pause_reason(t)):
                        stopped = paused
                        done = set(tcs)
                        self._pause(paused, [x for tg in order for x in plan[tg] if x not in done])
            for n, t in tcs.items():
                t.ctx = ""
                t.finish()
        finally:
            if restore:
                try:
                    session.switch_role(browser.page(), "home")
                except Exception:
                    traceback.print_exc(); print("!! role not restored:", session.role_command("home"))
                try:
                    print("env restore:", session.restore_env(orig))
                except Exception:
                    traceback.print_exc(); print("!! environment not restored; original state:", orig)
            try:
                if self.teardown_fn:
                    self.teardown_fn(self)
            except Exception:
                traceback.print_exc(); print("!! teardown failed: run tools.cleanup")
            browser.stop()

    def run_dev(self, nos, env, role):
        """Development: run the cases in the CURRENT state of the system (no switching, no setup / teardown)."""
        session._state.update(env=env, role=role)
        try:
            if not self._session_or_pending(nos):
                return
            for n in nos:
                t = self.new_tc(n)
                t.ctx = _label(env, role)
                t0 = datetime.datetime.now()
                self._run_case_settled(t, n, role, env)
                t1 = datetime.datetime.now()
                t.timings.append(dict(slot=_label(env, role).strip(" []") or "default", start=t0.strftime("%H:%M:%S"), end=t1.strftime("%H:%M:%S"),
                                      seconds=round((t1 - t0).total_seconds(), 1)))
                t.ctx = ""
                t.finish()
                if (reason := self._stop_reason(t)):
                    self._stop(reason, nos[nos.index(n) + 1:])
                    break
                if (paused := self._pause_reason(t)):
                    self._pause(paused, nos[nos.index(n) + 1:])
                    break
        finally:
            browser.stop()

    def _matrix_fn(self, n):
        import inspect
        return len(inspect.signature(self.cases[n]).parameters) >= 3

    def _record_partial(self, n, t, status, remark):
        results.record(n, title=t.title if t else str(self.spec.get(n, {}).get("view", "")), status=status, remark=remark, steps=t.steps if t else [],
                       started=t.started if t else "", finished=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       expected_steps=t.steps_text if t else [])
        print(f"== {n} => {status}: {remark}")

    def main(self, argv=None):
        argv = argv if argv is not None else sys.argv[1:]
        cfg.ensure()
        matrix = "--matrix" in argv
        only = [a.split("=", 1)[1] for a in argv if a.startswith("--targets=")]
        dev = [a.split("=", 1)[1] for a in argv if a.startswith("--dev=")]
        no_restore = "--no-restore" in argv
        ids = [a for a in argv if not a.startswith("--")]
        nos = [int(a) if a.isdigit() else a for a in ids] if ids else sorted(self.cases, key=str)
        unknown = [n for n in nos if n not in self.cases]
        if unknown:
            sys.exit(f"not registered in suite {self.name}: {unknown}")
        if dev:
            env, role = dev[0].split(":")
            return self.run_dev(nos, env, role)
        if matrix:
            ot = {tuple(x.split(":")) for x in only[0].split(",")} if only else None
            self.run_matrix(nos, only_targets=ot, restore=not no_restore)
        else:
            self.run(nos)
