"""Case scripts: one small Python file per case, kept in the workspace (scripts/<workbook name>/<case>.py) so the NEXT session that tests the same file
reuses them instead of paying for them again. `scaffold` writes the skeleton from the analysis (captions, expected results, data fallbacks and edge
cases are already in it); the assistant only fills in the actions. A script carries the digest of the case it was written for, so a changed case
is flagged and an unchanged one is reused untouched."""
import ast
import hashlib
import re
from pathlib import Path

from . import analysis, paths
from .state import Session
from .store import read_json

HEADER = re.compile(r"^# oqa-case: (\S+) sha=(\w+)", re.M)


def script_dir(session: Session) -> Path:
    stem = re.sub(r"[^\w.-]+", "_", Path(session.state()["file"]).stem) or "workbook"
    return paths.workspace_dir(session.ws) / "scripts" / stem


def script_path(session: Session, case) -> Path:
    return script_dir(session) / (re.sub(r"[^\w.-]+", "_", case) + ".py")


def case_sha(case_analysis):
    """Digest of what a script must follow: the case's steps with their expected results, its data fallbacks and what it depends on."""
    return analysis.digest({k: case_analysis.get(k) for k in ("steps", "data_needs", "depends_on", "profile")})


def unwritten(path: Path):
    """How many steps of the script still say `s.todo()`."""
    return path.read_text(encoding="utf-8").count(".todo(") if path.exists() else 0


CHECK_CALLS = ("check", "equals")


def script_sha(path: Path):
    """Digest of a case script's text: tells a repair that changed the script from one that only re-ran it."""
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16] if path.exists() else ""


STEP_FUNCTION = re.compile(r"step_(\d+)$")


def _parse(path: Path):
    try:
        return ast.parse(path.read_text(encoding="utf-8")) if path.exists() else None
    except SyntaxError:
        return None                     # the runner reports a syntax error with its line


def _blocks(tree):
    """(step number, the AST node holding its statements) for every step of a script, in either shape: `def step_<n>(t, s)` or the older `with t.step(<n>) as s:`."""
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and STEP_FUNCTION.match(node.name):
            yield int(STEP_FUNCTION.match(node.name).group(1)), node
        elif isinstance(node, ast.With):
            for item in node.items:
                call = item.context_expr
                if (isinstance(call, ast.Call) and isinstance(call.func, ast.Attribute) and call.func.attr == "step" and call.args
                        and isinstance(call.args[0], ast.Constant)):
                    yield call.args[0].value, node


def step_hashes(path: Path):
    """{step n: digest of the step's code} for a script written as step functions; {} for the older shape (a script that cannot resume).
    A step that did not change since it ran is one a repair of a later step does not have to run again."""
    tree = _parse(path)
    if tree is None:
        return {}
    text = path.read_text(encoding="utf-8")
    return {n: hashlib.sha256((ast.get_source_segment(text, node) or "").encode("utf-8")).hexdigest()[:12]
            for n, node in _blocks(tree) if isinstance(node, ast.FunctionDef)}


def _steps_lacking(path: Path, present):
    """Numbers of the steps in which no statement satisfies `present(node)`."""
    tree = _parse(path)
    if tree is None:
        return []
    return sorted(n for n, node in _blocks(tree) if not any(present(x) for stmt in node.body for x in ast.walk(stmt)))


def _is_check(x):
    return isinstance(x, ast.Assert) or (isinstance(x, ast.Call) and isinstance(x.func, ast.Attribute) and (x.func.attr in CHECK_CALLS or x.func.attr.startswith("expect_")))


def _is_highlight(x):
    return isinstance(x, ast.Call) and isinstance(x.func, ast.Attribute) and x.func.attr == "hl" and bool(x.args)


def unchecked_steps(path: Path):
    """Numbers of the `with t.step(n)` blocks that contain no check at all: no assert, no s.check / s.equals, no o.expect_*. Such a step could only ever 'pass' by not failing."""
    return _steps_lacking(path, _is_check)


def unhighlighted_steps(path: Path):
    """Numbers of the `with t.step(n)` blocks that never call s.hl(<selector>): the picture would have nothing to frame, so the step could not pass."""
    return _steps_lacking(path, _is_highlight)


def header(info, item, sha):
    """The comment lines at the top of a case script: which case and which version of it, then what the analysis said about it."""
    lines = [f"# oqa-case: {info['id']} sha={sha}", f"# {info['title']}", f"# profile: {item['profile']}"]
    if item.get("depends_on"):
        lines.append(f"# starts from the end state of: {', '.join(item['depends_on'])}" + (f" ({item['reuse']})" if item.get("reuse") else ""))
    lines += [f"# edge case: {e}" for e in item.get("edge_cases") or []]
    lines += [f"# data may be missing: {d['what']} -> if so: {d['fallback']}   (o.reach(enough, make) creates records named o.name(...))" for d in item.get("data_needs") or []]
    lines += ["# o = Odoo helpers: menu button field row rows pager toasts new fill m2o save search open_row expect_text expect_visible expect_value expect_rows expect_no_errors reach name unique",
              "# One function per step: step_<n>(t, s). A failed step is repaired and run again ALONE, from the page where it stopped, so steps share values through t.data,",
              "# never through local variables or module globals (t.data[\"order\"] = o.name(\"order\") in one step, t.data[\"order\"] in a later one).",
              "# Look at the page with `oqa probe` (or `oqa probe --failed` for the screen where a step stopped) when a selector does not fit."]
    return lines


def render(info, item, sha):
    lines = header(info, item, sha) + ["", ""]
    for step in item["steps"]:
        lines += [f"def step_{step['n']}(t, s):", "    o = t.o", "    page = t.page", f"    # {step['n']}. {step['action']}   ->   {step['expected']}"]
        lines += [
                  "    # s.hl(\"<selector>\")   name the element this step is about FIRST: a failing step is framed too",
                  "    # s.check(<what the page shows that proves the expected result>)   every step needs at least one check (s.check, s.equals, o.expect_*, assert)",
                  "    s.todo()", "", ""]
    return "\n".join(lines).rstrip("\n") + "\n"


def scaffold(session: Session, ids=None):
    """[{case, path, status}] where status is new (written now), reuse (same case, script untouched) or changed (the case changed since the script was written:
    review it against the new steps)."""
    session.require("EXECUTE", "CONFIRM", "CONNECT", "ANALYZE", working=False)
    cases = {c["id"]: c for c in session.cases()["cases"]}
    items = {c["id"]: c for c in (read_json(session.dir / "analysis.json") or {"cases": []})["cases"]}
    order = (read_json(session.dir / "strategy.json") or {}).get("order") or list(items)
    out = []
    for case in ids or order:
        if case not in items:
            out.append({"case": case, "path": "", "status": "unknown"})
            continue
        path, sha = script_path(session, case), case_sha(items[case])
        existing = HEADER.search(path.read_text(encoding="utf-8")) if path.exists() else None
        if existing is None:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(render(cases[case], items[case], sha), encoding="utf-8")
            out.append({"case": case, "path": str(path), "status": "new"})
        else:
            out.append({"case": case, "path": str(path), "status": "reuse" if existing.group(2) == sha else "changed", "unwritten": unwritten(path),
                        "unchecked": unchecked_steps(path), "unhighlighted": unhighlighted_steps(path)})
    return out
