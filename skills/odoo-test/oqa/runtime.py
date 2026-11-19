"""EXECUTE: run a case script and make the evidence hard to fake.

A script is `def run(t)` with one `with t.step(n) as s:` block per step of the analysis. The runner takes the caption and the expected result from the
analysis (the script never repeats them), photographs EVERY step, PASS or FAIL, frames the element named by `s.hl(...)`, and refuses to call a case
PASS unless every step ran in order, passed, and has a picture pair whose red frame is really there."""
import ast
import contextlib
import re
import textwrap
import time
import types
from pathlib import Path

from . import connect, diagnose, evidence, events, knowledge, memory, probe, recipe, refs
from .browser import launch
from .odoo_ui import Odoo
from .scaffold import HEADER, STEP_FUNCTION, case_sha, script_path, script_sha, step_hashes, unchecked_steps, unhighlighted_steps, unwritten
from .state import MAX_REPAIRS, GateError, Session
from .store import SchemaError, read_json


class StepFailed(Exception):
    pass


class UserStopped(Exception):
    """The user pressed Stop: the step in progress was finished, and no further step starts."""


def reason_of(error):
    """One line that says what went wrong, with the element Playwright was waiting for when it names one."""
    lines = str(error).splitlines()
    reason = lines[0][:300] if lines else type(error).__name__
    if not isinstance(error, (StepFailed, AssertionError)):
        reason = f"{type(error).__name__}: {reason}"
        waited = next((line.strip() for line in lines[1:] if "waiting for" in line), "")          # Playwright's call log names the element it was waiting for
        reason += f" ({waited[:140]})" if waited else ""
    return reason


def _clip(box, width, height):
    x0, y0, x1, y1 = max(box["x"], 0), max(box["y"], 0), min(box["x"] + box["width"], width), min(box["y"] + box["height"], height)
    return (x0, y0, x1 - x0, y1 - y0) if x1 - x0 >= 2 and y1 - y0 >= 2 else None


class Step:
    """`with t.step(n) as s:` — `s.hl(...)` names what the picture is about; leaving the block photographs it and records PASS or FAIL."""

    def __init__(self, t, n):
        self.t, self.n, self.targets, self.note = t, n, [], ""
        spec = t.spec[n - 1]
        self.caption, self.expected = spec["action"], spec["expected"]

    def hl(self, *targets):
        self.targets += [x for x in targets if x is not None]
        return self

    def actual(self, text):
        self.note = text

    def todo(self):
        raise StepFailed(f"step {self.n} is not written yet: fill in the script")

    def check(self, condition, message=""):
        """A check of the step's expected result. A step with no check cannot pass: a picture and no error prove nothing about what the page shows."""
        self.t.checks += 1
        if not condition:
            raise StepFailed(message or "a check of this step failed")

    def equals(self, actual, expected, what="value"):
        self.check(actual == expected, f"{what} is {actual!r}, expected {expected!r}")

    def __enter__(self):
        t = self.t
        if t.session.state().get("stopped"):
            raise UserStopped()
        if self.n != len(t.done) + 1:
            raise StepFailed(f"steps must run in order: expected step {len(t.done) + 1}, got {self.n}")
        self.checks_before = t.checks
        return self

    def _boxes(self, failed):
        page = self.t.page
        size = page.viewport_size
        boxes = []
        for target in self.targets:
            loc = page.locator(target) if isinstance(target, str) else target
            try:
                loc.first.scroll_into_view_if_needed(timeout=2000)
                found = [b for b in (loc.nth(i).bounding_box() for i in range(min(loc.count(), 30))) if b]
            except Exception:
                found = []
            if found:
                union = {"x": min(b["x"] for b in found), "y": min(b["y"] for b in found)}
                union["width"] = max(b["x"] + b["width"] for b in found) - union["x"]
                union["height"] = max(b["y"] + b["height"] for b in found) - union["y"]
                box = _clip(union, size["width"], size["height"])
                if box:
                    boxes.append(box)
        if boxes or not failed:
            return boxes
        try:                                        # a FAIL still gets a frame: the error the page shows, else the whole screen
            loc = page.locator(".o_notification.border-danger, .modal:visible, .alert-danger, .o_field_invalid")
            box = _clip(loc.first.bounding_box(), size["width"], size["height"]) if loc.count() else None
        except Exception:
            box = None
        return [box or (0, 0, size["width"], size["height"])]

    def _capture(self, status):
        t = self.t
        t.page.wait_for_timeout(250)
        folder = t.evidence_dir
        folder.mkdir(parents=True, exist_ok=True)
        raw, annotated = folder / f"s{self.n}_raw.png", folder / f"s{self.n}_annotated.png"
        boxes = self._boxes(failed=status == "FAIL")
        t.page.screenshot(path=str(raw), animations="disabled")
        evidence.annotate(raw, boxes, f"Step {self.n}/{len(t.spec)} {status}: {self.caption}", annotated, ok=status == "PASS")
        return {"raw": raw, "annotated": annotated, "boxed": bool(self.targets) and bool(boxes)}

    def __exit__(self, kind, error, trace):
        t = self.t
        reason = reason_of(error) if error is not None else None
        status = "PASS" if error is None else "FAIL"
        pictures, diag = None, None
        try:
            pictures = self._capture(status)
            if status == "PASS" and not pictures["boxed"]:
                status, reason = "FAIL", "no element was boxed for this step: call s.hl(<the element this step is about>)"
            elif status == "PASS":
                problem = evidence.check(pictures["raw"], pictures["annotated"])
                if problem:
                    status, reason = "FAIL", problem
            if status == "FAIL":
                diag = diagnose.read(t.page)
        except Exception as capture_error:
            if status == "PASS":
                status, reason = "FAIL", f"could not take the evidence: {str(capture_error)[:200]}"
        checks = t.checks - self.checks_before
        if status == "PASS" and not checks:
            status, reason = "FAIL", "no check ran in this step: compare what the page shows with the expected result (s.check, s.equals, o.expect_* or assert)"
            diag = diagnose.read(t.page)
        if status == "FAIL":
            try:
                probe.save_page(t.page, t.evidence_dir / probe.FAILED_FILE)          # the screen the step stopped on: `oqa probe --failed` reads it without a browser
            except Exception:
                pass
        actual = self.note or "as expected" if status == "PASS" else f"{reason}" + (f" | page says: {diag['summary']}" if diag else "")
        rel = (lambda p: Path(p).relative_to(t.session.dir).as_posix())
        try:
            events.append(t.session.dir, "step", case=t.case, attempt=t.attempt, n=self.n, status=status, caption=self.caption, expected=self.expected,
                          actual=actual, checks=checks, images=[{"raw": rel(pictures["raw"]), "annotated": rel(pictures["annotated"])}] if pictures else [])
        except SchemaError:
            pass                                    # no picture could be taken: the attempt result below says so
        t.done.append({"n": self.n, "status": status, "reason": reason, "diag": diag, "pictures": pictures})
        if status == "FAIL":
            raise StepFailed(reason or "failed") from error
        return False


class Trial(Step):
    """The step object while a piece of a step is TRIED on the real page: the same methods, but `s.hl(...)` has to find its element on the page, now."""

    def hl(self, *targets):
        for target in targets:
            if target is None:
                continue
            locator = self.t.page.locator(target) if isinstance(target, str) else target
            try:
                found = locator.count() > 0 and locator.first.is_visible()
            except Exception:
                found = False
            if not found:
                raise StepFailed(f"s.hl: nothing visible on the page matches {target if isinstance(target, str) else 'that locator'}")
        return super().hl(*targets)


class T:
    """What a case script receives."""

    def __init__(self, runner, case, attempt, spec, profile, page, timeout=None, data=None, done=None):
        self.runner, self.session, self.case, self.attempt, self.spec, self.profile, self.page = runner, runner.session, case, attempt, spec, profile, page
        self.checks = 0
        self.o = Odoo(page, runner.ws["url"], runner.ws["test_prefix"], on_check=self._count, timeout=timeout)
        self.o.resolve_ref = refs.resolver(runner.session, runner.ws)
        self.evidence_dir = runner.session.dir / "evidence" / re.sub(r"[^\w.-]+", "_", case) / f"a{attempt}"
        self.done = list(done or [])                 # steps already passed (carried over when this run picks up where the last one stopped)
        self.data = {} if data is None else data     # what steps hand to each other; it outlives a failed step, so a repaired step finds what the earlier ones made

    def _count(self):
        self.checks += 1

    def step(self, n):
        if not 1 <= n <= len(self.spec):
            raise StepFailed(f"the analysis has {len(self.spec)} steps for this case; there is no step {n}")
        return Step(self, n)


_NO_ARGS = ast.arguments(posonlyargs=[], args=[], kwonlyargs=[], kw_defaults=[], defaults=[])


class _CountedAsserts(ast.NodeTransformer):
    """`assert cond, message` in a case script becomes a call that counts as a check and raises the same AssertionError; the message is only built on failure."""

    def visit_Assert(self, node):
        lazy = lambda body: ast.Lambda(args=_NO_ARGS, body=body)
        call = ast.Call(func=ast.Name("_oqa_assert", ast.Load()), args=[lazy(node.test), lazy(node.msg or ast.Constant(""))], keywords=[])
        return ast.copy_location(ast.Expr(call), node)


def _load(path, count):
    """Run the script's module code with its asserts counted, and return the module."""
    tree = ast.fix_missing_locations(_CountedAsserts().visit(ast.parse(path.read_text(encoding="utf-8"), str(path))))

    def counted_assert(test, message):
        count()
        if not test():
            raise AssertionError(message())

    module = types.ModuleType("oqa_case_" + re.sub(r"\W", "_", path.stem))
    module.__file__ = str(path)
    module._oqa_assert = counted_assert
    exec(compile(tree, str(path), "exec"), module.__dict__)
    return module


def _step_functions(module, path, total):
    """{n: function} for a script written as `def step_<n>(t, s)`; {} for the older `def run(t)` (which cannot pick up from a step)."""
    found = {}
    for name, function in vars(module).items():
        match = STEP_FUNCTION.match(name)
        if match and callable(function):
            found[int(match.group(1))] = function
    if found:
        missing = [n for n in range(1, total + 1) if n not in found]
        if missing:
            raise StepFailed(f"{path.name} has no function for step(s) {missing}: one `def step_<n>(t, s)` per step of the analysis")
        return found
    if not callable(getattr(module, "run", None)):
        raise StepFailed(f"{path.name} has no `def step_<n>(t, s)` for its steps and no `def run(t)`")
    return {}


class Runner:
    """Runs cases. One Runner that lives across several runs (the runner service) also keeps each case's browser page, the steps it passed and the data they made, so a
    repaired step runs again ALONE, from the page where it stopped, instead of the whole case from step 1."""

    def __init__(self, session: Session, ws, store, headless=True):
        self.session, self.ws, self.store, self.headless = session, ws, store, headless
        self.analysis = {c["id"]: c for c in read_json(session.dir / "analysis.json")["cases"]}
        self.pages, self.last, self.last_profile = {}, None, None
        self.memory = {}

    def _page(self, new_page, case_analysis, keep_screen=False):
        profile = next(p for p in self.ws["profiles"] if p["name"] == case_analysis["profile"])
        if profile["name"] not in self.pages:
            self.pages[profile["name"]] = new_page(connect.storage_path(self.session, profile["name"]))[1]
        page = self.pages[profile["name"]]
        continuing = self.last in (case_analysis.get("depends_on") or []) and self.last_profile == profile["name"]
        if not (continuing or keep_screen):          # a case that builds on the one before, or a step picked up where it stopped, keeps its screen; any other starts from the landing page
            landing = ((self.session.state().get("connect") or {}).get("profiles", {}).get(profile["name"]) or {}).get("landing_url") or self.ws["url"] + "/web"
            page.goto(landing, wait_until="domcontentloaded")
            if "/web/login" in page.url:
                connect.login(page, self.ws, profile, self.store)
            page.wait_for_selector(".o_web_client", state="visible", timeout=15000)
        return profile, page

    # --- exploring: try the code of a step on the real page BEFORE it is written into a script -----------------------------
    def _timeout(self, case, item):
        return ((self.session.state().get("timeouts") or {}).get(case) or item.get("timeout_s") or 0) * 1000 or None

    def _env(self, ex, step):
        """The names a piece of step code sees while it is tried (the same as in a script), kept for the whole step so a piece can use what an earlier piece made."""
        t = ex["t"]

        def counted(test, message):
            t.checks += 1
            if not test():
                raise AssertionError(message())
        return ex["envs"].setdefault(step, {"t": t, "o": t.o, "s": Trial(t, step), "page": t.page, "_oqa_assert": counted})

    @staticmethod
    def _exec(env, code):
        tree = ast.fix_missing_locations(_CountedAsserts().visit(ast.parse(textwrap.dedent(code), "<try>")))
        exec(compile(tree, "<try>", "exec"), env)

    def _look(self, ex, needle=""):
        """What the page shows now, in few lines: where it is, what it says, and the controls (all of them when filtered, the first few otherwise)."""
        page = ex["t"].page
        try:
            ex["t"].o.ready()                       # a page still loading would show missing controls: look only once it has settled
            diag = diagnose.read(page)
            return {"url": page.url, "title": page.title(), "toasts": ex["t"].o.toasts()[:3], "page_says": diag["summary"],
                    "controls": page.evaluate(probe.JS, [needle, probe.SHOWN if needle else 25])}
        except Exception as error:
            return {"url": "", "title": "", "toasts": [], "page_says": f"could not read the page: {str(error)[:100]}", "controls": []}

    def explore(self, new_page, case, to=1, needle=""):
        """Start exploring `case`: a new page at the landing screen, then every verified piece of the steps before `to` run again, so the page is where step `to` begins."""
        item = self.analysis.get(case)
        if item is None:
            raise GateError(f"{case} is not in the analysis.")
        profile, page = self._page(new_page, item)
        self.last, self.last_profile = None, None            # the page moves from here: a failed step can no longer be picked up where it stopped
        self.memory.pop(case, None)
        t = T(self, case, 0, item["steps"], profile["name"], page, self._timeout(case, item), data={})
        ex = self.exploring = {"case": case, "t": t, "envs": {}, "step": 0}
        done = recipe.steps_of(recipe.load(self.session, case))
        if to > 1:
            events.append(self.session.dir, "try_started", case=case, step=to, mode="replay")      # taking the page to step `to` by replaying the verified pieces
        for n in range(1, to):
            if n not in done:
                raise GateError(f"step {n} has nothing verified yet, so the page cannot be taken to step {to}: try step {n} first (oqa try {case} --step {n} --code ...)")
            for piece in done[n]:
                try:
                    self._exec(self._env(ex, n), piece)
                except Exception as error:
                    events.append(self.session.dir, "explore", case=case, step=n, ok=False, error=reason_of(error)[:200], url=page.url)
                    return {"ok": False, "step": n, "error": f"a verified piece of step {n} no longer works: {reason_of(error)}", "code": piece, **self._look(ex, needle)}
            ex["step"] = n
        events.append(self.session.dir, "explore", case=case, step=to, ok=True, url=page.url)
        return {"ok": True, "step": ex["step"], "at": to, **self._look(ex, needle)}

    def try_step(self, case, step, code="", needle="", reset=False, undo=False):
        """Run `code` for `step` on the page the exploration is at. When it works it is kept as the next verified piece of the step; when it does not, nothing is kept
        and no attempt is used: this is where selectors are found out, not in a failed run."""
        ex = getattr(self, "exploring", None)
        if not ex or ex["case"] != case:
            raise GateError(f"Not exploring {case}: run `oqa explore {case}` first (it opens the page where step 1 starts).")
        item = self.analysis[case]
        if not 1 <= step <= len(item["steps"]):
            raise GateError(f"{case} has {len(item['steps'])} steps; there is no step {step}.")
        if undo:
            pieces = recipe.record(self.session, case, step, "", undo=True)
            return {"ok": True, "step": step, "pieces": len(pieces), "undone": True, **self._look(ex, needle)}
        if step > ex["step"] + 1:
            raise GateError(f"The page is at step {ex['step']}: try step {ex['step'] + 1} before step {step}.")
        if step < ex["step"]:
            raise GateError(f"The page is already past step {step}: go back with `oqa explore {case} --to {step}` (it replays what is verified).")
        t, started = ex["t"], time.monotonic()
        t.checks = 0
        events.append(self.session.dir, "try_started", case=case, step=step)          # the dashboard shows "trying step N" only between this and the explore event below
        try:
            self._exec(self._env(ex, step), code)
        except Exception as error:
            events.append(self.session.dir, "explore", case=case, step=step, ok=False, code=code[:300], error=reason_of(error)[:200], url=t.page.url)
            look = self._look(ex, needle)
            try:
                everything = t.page.evaluate(probe.JS, ["", probe.KEPT])
            except Exception:
                everything = []
            return {"ok": False, "step": step, "error": reason_of(error), "seconds": round(time.monotonic() - started, 1), "similar": probe.near(everything, reason_of(error)), **look}
        pieces = recipe.record(self.session, case, step, code, reset=reset)
        ex["step"] = max(ex["step"], step)
        events.append(self.session.dir, "explore", case=case, step=step, ok=True, code=code[:300], url=t.page.url)
        return {"ok": True, "step": step, "pieces": len(pieces), "checks": t.checks, "seconds": round(time.monotonic() - started, 1), **self._look(ex, needle)}

    def _begin(self, case, item, path, start, fresh):
        """(step to start at, what is kept, a note for the assistant). A run picks up at the step that failed when the page is still where that step left it: the same
        case ran last, on the same profile, in this browser, and none of the steps before it was edited since it passed."""
        hashes = step_hashes(path)
        mem = self.memory.get(case)
        usable = bool(mem and mem["failed"] and hashes and self.last == case and self.last_profile == item["profile"])
        edited = [n for n in range(1, mem["failed"]) if hashes.get(n) != mem["hashes"].get(n)] if usable else []
        if start and start > 1:
            if not usable or edited or fresh or start > mem["failed"]:
                raise GateError(f"{case} cannot start at step {start}: " + ("step(s) " + ", ".join(map(str, edited)) + " were edited after they passed; run without --from to redo them" if edited
                                                                         else "no earlier run of it is kept in the runner's browser (or it stopped before that step). Run it from step 1 (--fresh)"))
            begin, note = start, f"starting at step {start} as asked"
        elif usable and not edited and not fresh:
            begin, note = mem["failed"], f"picking up at step {mem['failed']}: steps 1-{mem['failed'] - 1} passed before and are kept"
        else:
            begin = 1
            note = (f"step(s) {', '.join(map(str, edited))} were edited after they passed: starting from step 1" if edited else "")
        kept = {"done": list(mem["done"][:begin - 1]), "data": mem["data"]} if begin > 1 else {"done": [], "data": {}}
        return begin, kept, note, hashes

    def run_case(self, new_page, case, start=None, fresh=False):
        """One attempt at one case. Returns a compact result dict; the case is finished (PASS) only when every hard check holds."""
        item = self.analysis.get(case)
        if item is None:
            raise GateError(f"{case} is not in the analysis.")
        path = script_path(self.session, case)
        if not path.exists():
            raise GateError(f"No script for {case}: run oqa scaffold {case}, then fill it in.")
        if self.session.state().get("explore") == "required":
            recipe.require_sealed(path, case)
        if unwritten(path):
            raise GateError(f"{path.name} still has {unwritten(path)} unwritten step(s) (s.todo()). Fill them in first; this does not use up an attempt.")
        bare = unchecked_steps(path)
        if bare:
            raise GateError(f"{path.name}: step(s) {bare} have no check. Each step must compare what the page shows with its expected result: s.check(...), s.equals(...), "
                            "o.expect_* or assert. Add one; this does not use up an attempt.")
        loose = unhighlighted_steps(path)
        if loose:
            raise GateError(f"{path.name}: step(s) {loose} never call s.hl(<selector>). The picture of a step frames the element the step is about; name it first in the block "
                            "(probe the page if you do not know the selector). Add it; this does not use up an attempt.")
        header = HEADER.search(path.read_text(encoding="utf-8"))
        stale = bool(header and header.group(2) != case_sha(item))
        begin, kept, note, hashes = self._begin(case, item, path, start, fresh)
        attempt = self.session.begin_attempt(case)
        sha = script_sha(path)
        timeout = ((self.session.state().get("timeouts") or {}).get(case) or item.get("timeout_s") or 0) * 1000 or None
        events.append(self.session.dir, "case_started", case=case, attempt=attempt, **({"from_step": begin} if begin > 1 else {}))
        profile, page = self._page(new_page, item, keep_screen=begin > 1)
        t = T(self, case, attempt, item["steps"], profile["name"], page, timeout, data=kept["data"], done=kept["done"])
        failure, stopped = None, False
        try:
            module = _load(path, t._count)
            functions = _step_functions(module, path, len(item["steps"]))
            if functions:
                for n in range(begin, len(item["steps"]) + 1):
                    with t.step(n) as s:
                        functions[n](t, s)
            else:
                module.run(t)
        except UserStopped:
            stopped = True
        except Exception as error:
            failure = error
        hints = [h["text"] for h in memory.hints(self.ws["name"], memory.stem_of(self.session.state()["file"]), case)][-3:]       # this session's and earlier sessions' alike
        passed = [d for d in t.done if d["status"] == "PASS"]
        self.memory[case] = {"done": passed, "data": t.data, "failed": len(passed) + 1 if (stopped or failure is not None) and hashes else None, "hashes": hashes}
        if stopped:                                  # the user's Stop is not the case's failure: it costs no attempt and leaves no FAIL step
            self.session.refund_attempt(case)
            events.append(self.session.dir, "attempt_finished", case=case, attempt=attempt, status="STOPPED", reason="stopped by the user", script_sha=sha)
            self.last, self.last_profile = (case, profile["name"]) if hashes else (None, None)
            return {"case": case, "attempt": attempt, "status": "STOPPED", "steps_passed": len(passed), "steps_total": len(item["steps"]), "from_step": begin, "note": note,
                    "failed_step": None, "reason": "stopped by the user", "page_says": "", "stale_script": stale, "knowledge": [], "picture": "", "left": self._left(), "hints": hints,
                    "similar": [], "failed_page": "", "runs_left": self._runs_left(case)}
        if failure is None and len(t.done) < len(item["steps"]):
            failure = StepFailed(f"the script ended after step {len(t.done)} of {len(item['steps'])}")
        if failure is not None and not any(d["status"] == "FAIL" for d in t.done):
            failure = self._unrecorded(t, failure)
            passed = [d for d in t.done if d["status"] == "PASS"]
            self.memory[case]["done"], self.memory[case]["failed"] = passed, (len(passed) + 1 if hashes else None)
        if failure is None:                          # the hard check, again, over everything this attempt wrote
            for done in t.done:
                problem = evidence.check(done["pictures"]["raw"], done["pictures"]["annotated"]) if done["pictures"] else "no picture"
                if problem or done["status"] != "PASS":
                    failure = StepFailed(f"step {done['n']}: {problem or done['reason']}")
                    break
            if failure is not None:
                self.memory[case]["failed"] = None   # evidence that no longer holds: only a run from step 1 can make it again
        failed = next((d for d in t.done if d["status"] == "FAIL"), None)
        reason = str(failure) if failure is not None else ""
        status = "FAIL" if failure is not None else "PASS"
        events.append(self.session.dir, "attempt_finished", case=case, attempt=attempt, status=status, reason=reason, script_sha=sha, **({"failed_step": failed["n"]} if failed else {}))
        if status == "PASS":
            self.session.finish_case(case, "PASS")
        elif failed:
            self.session.advance(case, failed["n"])  # got further than before: this run was not a failed repair
        self.last, self.last_profile = case, profile["name"]
        page_says = failed["diag"]["summary"] if failed and failed["diag"] else ""
        related = knowledge.related(self.ws["name"], reason + " " + page_says) if status == "FAIL" else []
        saved = t.evidence_dir / probe.FAILED_FILE
        similar = probe.near(probe._read(saved)[2], reason + " " + page_says) if status == "FAIL" and saved.exists() else []
        return {"case": case, "attempt": attempt, "status": status, "steps_passed": len(passed), "steps_total": len(item["steps"]), "from_step": begin, "note": note,
                "failed_step": failed["n"] if failed else None, "reason": reason, "page_says": page_says, "stale_script": stale, "knowledge": related,
                "picture": str(failed["pictures"]["annotated"]) if failed and failed["pictures"] else "", "left": self._left(), "hints": hints,
                "similar": similar, "failed_page": str(saved) if status == "FAIL" and saved.exists() else "", "runs_left": self._runs_left(case)}

    def _unrecorded(self, t, failure):
        """The script broke between steps (outside any step): photograph the screen as the failure of the step that should have come next."""
        n = len(t.done) + 1
        if n > len(t.spec):
            return failure
        step = Step(t, n)
        try:
            with step:
                raise failure
        except StepFailed as recorded:
            return recorded

    def _runs_left(self, case):
        """Runs the case still has before the assistant must ask the user (the first run plus MAX_REPAIRS repairs, counted from the last fresh budget)."""
        return max(0, 1 + MAX_REPAIRS - self.session.state()["attempts"].get(case, 0))

    def _left(self):
        results = self.session.state()["results"]
        return [c for c in read_json(self.session.dir / "strategy.json")["order"] if c not in results]

    def run(self, cases, start=None, fresh=False, new_page=None):
        """Run cases one after another in ONE browser (the caller's, when it keeps one open), so a case can start where the one before ended. Stops at the first case that
        fails: that one needs repair. `start` (the first case only) picks the step to begin at; `fresh` forgets what an earlier run kept."""
        out = []
        with (contextlib.nullcontext(new_page) if new_page else launch(self.headless)) as pages:
            for i, case in enumerate(cases):
                if self.session.state().get("stopped"):
                    break
                out.append(self.run_case(pages, case, start if i == 0 else None, fresh))
                if out[-1]["status"] != "PASS":
                    break
        return out
