"""The recipe of a case: for each step, the code that was RUN on the real page and worked (`oqa try`), in order. The script is built from it (`oqa build`), so a line of a
script is never a guess about a selector: it ran once, on this Odoo, and did what the step says, before it was written down.

    <home>/workspaces/<ws>/scripts/<workbook>/<case>.recipe.json     {"schema_version": 1, "steps": {"1": ["code", ...], "2": [...]}}

It lives beside the script (not in the session), so the next session that tests the same workbook starts from steps already proven."""
import hashlib
import textwrap

from . import SCHEMA_VERSION
from .scaffold import script_path
from .state import GateError, Session
from .store import read_json, write_json

MARK = "# oqa-built: "


def path(session: Session, case):
    return script_path(session, case).with_suffix(".recipe.json")


def load(session: Session, case):
    data = read_json(path(session, case), None) or {"schema_version": SCHEMA_VERSION, "steps": {}}
    data.setdefault("steps", {})
    return data


def steps_of(data):
    return {int(n): list(code) for n, code in data["steps"].items() if code}


def record(session: Session, case, step, code, reset=False, undo=False):
    """Keep `code` as the next verified piece of `step` (after it ran well); `reset` starts the step over, `undo` drops its last piece."""
    data = load(session, case)
    pieces = [] if reset else list(data["steps"].get(str(step), []))
    if undo:
        pieces = pieces[:-1]
    else:
        pieces.append(textwrap.dedent(code).strip("\n"))
    data["steps"][str(step)] = pieces
    write_json(path(session, case), data)
    return pieces


def build(session: Session, case, item, title, header_lines):
    """The text of the script for `case`: one function per step of the analysis, made only of verified code (a step with none yet says so with s.todo())."""
    data = steps_of(load(session, case))
    out = []
    for step in item["steps"]:
        n = step["n"]
        out += [f"def step_{n}(t, s):", "    o = t.o", "    page = t.page", f"    # {n}. {step['action']}   ->   {step['expected']}"]
        if n in data:
            for piece in data[n]:
                out += ["    " + line if line.strip() else "" for line in piece.splitlines()]
        else:
            out += ["    # not explored yet: oqa try " + case + f" --step {n} --code \"...\"", "    s.todo()"]
        out += ["", ""]
    body = "\n".join(header_lines + [""] + out).rstrip("\n") + "\n"
    return body


def seal(text):
    """`text` with the line that says it was built, and a digest of everything after it: a later hand edit shows as a different digest."""
    first, _, rest = text.partition("\n")
    return f"{first}\n{MARK}{hashlib.sha256(rest.encode('utf-8')).hexdigest()[:16]}\n{rest}"


def sealed(script_text):
    """True when the script is exactly what `build` wrote: its mark matches its own content."""
    lines = script_text.split("\n", 2)
    if len(lines) < 3 or not lines[1].startswith(MARK):
        return False
    return lines[1][len(MARK):].strip() == hashlib.sha256(lines[2].encode("utf-8")).hexdigest()[:16]


def why_unsealed(script_text):
    lines = script_text.split("\n", 2)
    return "was edited by hand after it was built" if len(lines) > 1 and lines[1].startswith(MARK) else "was not built from steps tried on the real page"


def require_sealed(path_, case):
    text = path_.read_text(encoding="utf-8")
    if not sealed(text):
        raise GateError(f"{path_.name} {why_unsealed(text)}. A script is not written from memory of another Odoo: try each step on the real page, then build it. "
                        f"oqa explore {case} -> oqa try {case} --step N --code \"...\" (every step) -> oqa build {case}. This does not use up an attempt.")
