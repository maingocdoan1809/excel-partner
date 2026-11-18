# -*- coding: utf-8 -*-
"""Generate a SUITE SKELETON from a testcase workbook.
  python -m qakit.tools.gen_suite --spec testcases/<file>.xlsx --name <suite> [--ids ID,ID] [--force]
Creates suites/<name>/{__init__.py, cases.py, README.md}:
 * one function per testcase (matrix signature fn(t, role, env)) with the spec text in the docstring;
 * every spec step becomes `t.todo(...)` -> NOT EXECUTED until the agent replaces it with real actions + assertions + t.step(...), so an
   unwritten step can never become a PASS.
Then: write page objects in pages/, fill the cases, run `python -m suites.<name>.cases --dev=<env>:<role> <ID>` (no switching), then
`python -m qakit.tools.catalog build --note "automate ..."`."""
import argparse, re, sys
from pathlib import Path
from ..core.config import WORK, cfg
from ..reporting import spec as spec_mod

sys.stdout.reconfigure(encoding="utf-8")


def _fn(no):
    return "c_" + re.sub(r"\W", "_", str(no)).lower()


def build(spec_path, name, ids=None, force=False):
    sp = Path(spec_path) if Path(spec_path).is_absolute() else WORK / spec_path
    cases = [c for c in spec_mod.read(sp) if not ids or str(c["no"]) in ids]
    d = WORK / "suites" / name
    if (d / "cases.py").exists() and not force:
        sys.exit(f"{d / 'cases.py'} exists: add the cases by hand (or --force to overwrite)")
    d.mkdir(parents=True, exist_ok=True)
    (WORK / "suites" / "__init__.py").touch()
    (d / "__init__.py").write_text("", encoding="utf-8")
    rel = sp.relative_to(WORK).as_posix() if sp.is_relative_to(WORK) else str(sp)
    body = ['# -*- coding: utf-8 -*-',
            f'"""Suite {name} - generated from {sp.name}. Replace every t.todo() with real actions + assertions + t.step().',
            'Rules: change ONLY test objects (names start with the test prefix, guard.check); read data back to prove a write; no step without a screenshot."""',
            'from qakit.core import guard, session, ui',
            'from qakit.core.config import cfg',
            'from qakit.core.tc import Suite, Pending',
            '',
            f'SPEC = r"{rel}"',
            f'suite = Suite("{name}", SPEC)', '', '']
    for c in cases:
        no = c["no"]
        body += [f'@suite.case({no!r})', f'def {_fn(no)}(t, role, env):', f'    """{str(c.get("view", "")).strip()}',
                 f'    Role / env : {str(c.get("role", "")).strip()}',
                 f'    Precondition: {str(c.get("pre", "")).strip()}',
                 f'    Expected   : {str(c.get("exp", "")).strip()}', '    """',
                 '    pg = t.pg']
        steps = c["act_lines"] or [str(c.get("act", "")) or str(c.get("exp", ""))]
        for i, st in enumerate(steps, 1):
            body.append(f'    # ---- step {i}: {st}'.replace("\n", " "))
            cap = re.sub(r"^\s*\d+[.)．、]\s*", "", st)[:80].replace('"', "'").replace("\\", "/")
            body.append(f'    t.todo({i}, "{cap}", t.steps_text[{i - 1}] if len(t.steps_text) >= {i} else "")')
        body += ['', '']
    body += ['if __name__ == "__main__":', '    suite.main()', '']
    (d / "cases.py").write_text("\n".join(body), encoding="utf-8")
    (d / "README.md").write_text(
        f"# Suite {name}\nGenerated from `{sp.name}` ({len(cases)} cases).\n\n"
        "1. Read the screens (source + DOM: `python -m qakit.tools.dom_probe <path>`), write page objects in `pages/`.\n"
        f"2. Fill `cases.py`, try one case without switching: `python -m suites.{name}.cases --dev=<env>:<role> <ID>`.\n"
        "3. Register: `python -m qakit.tools.catalog build --note \"automate ...\"`, then `/system-test plan <IDs>`.\n", encoding="utf-8")
    print(f"generated {d} with {len(cases)} case(s)")
    return d


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--ids", help="only these IDs (comma separated)")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()
    build(a.spec, a.name, set(a.ids.split(",")) if a.ids else None, a.force)
