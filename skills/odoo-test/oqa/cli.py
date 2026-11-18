"""The `oqa` command line. Each command returns (data, text): `--json` prints the data, otherwise the text.

Exit codes: 0 done · 1 unexpected failure · 2 refused (wrong phase, invalid input, waiting for the user). A refusal always says what to run instead."""
import argparse
import json
import os
import shutil
import signal
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path
from urllib.parse import quote

from . import VERSION, build_id, analysis, connect, memory, dashboard_server, doctor, events, init_server, intake, knowledge, paths, probe, recipe, report, runtime, scaffold, service, workspace
from .secrets import SecretStoreUnavailable, default_store
from .state import MAX_REPAIRS, GateError, Session
from .store import SchemaError, read_json, write_json

REFUSED = (GateError, SchemaError, paths.PathError, knowledge.KnowledgeError, SecretStoreUnavailable, connect.LoginFailed)


# --- commands ---------------------------------------------------------------------
def cmd_init(args):
    if args.status:
        data = init_server.status(args.ws)
        return data, ("Workspace " + data["workspace"] + " is configured (" + data["url"] + ")") if data["configured"] else "No valid workspace yet. Run: oqa init"
    server = init_server.InitServer(default_store(), args.ws, args.timeout)
    summary = server.run(lambda url: print(f"Waiting for the form to be saved: {url}", flush=True), open_browser=not args.no_open)
    if summary is None:
        raise GateError(f"The form was not saved within {args.timeout} seconds. Nothing was written. Run oqa init again.")
    return summary, f"Saved workspace {summary['workspace']} ({summary['url']}), profiles: {', '.join(summary['profiles'])}"


def cmd_doctor(args):
    report = doctor.check()
    if args.fix and not report["ready"]:
        report = doctor.fix(report, lambda line: print(line, file=sys.stderr, flush=True))
    lines = [f"oqa {report['oqa']} · python {report['python']} {'ok' if report['python_ok'] else 'TOO OLD (3.9+ needed)'}",
             "packages: " + ("ok" if not report["missing_packages"] else "missing " + ", ".join(report["missing_packages"])),
             "browser: " + ("ok" if report["browser_installed"] else "not installed"),
             "password store: " + ("ok" if report["secret_store"] else report["secret_store_note"]),
             "READY" if report["ready"] else "Not ready. Run: oqa doctor --fix"]
    return report, "\n".join(lines)


def cmd_ws(args):
    if args.what == "list":
        names = paths.workspace_names()
        return {"workspaces": names}, "\n".join(names) or "No workspace yet. Run: oqa init"
    name = paths.resolve_workspace(args.ws)
    ws = workspace.load(name)
    try:
        view = workspace.public_view(ws, default_store())
    except SecretStoreUnavailable:
        view = ws
    lines = [f"{view['name']}  {view['url']}" + (f"  db={view['db']}" if view["db"] else "")]
    lines += [f"  profile {p['name']}: {p['login']}" + ("" if "has_secret" not in p else " (password saved)" if p["has_secret"] else " (NO PASSWORD)") + (f"  {p['note']}" if p["note"] else "") for p in view["profiles"]]
    lines += [f"  reference folder: {d}" for d in view["reference_dirs"]]
    return view, "\n".join(lines)


def cmd_session(args):
    if args.what == "list":
        ws = paths.resolve_workspace(args.ws)
        items = [Session(ws, i).summary() for i in paths.session_ids(ws)]
        return {"sessions": items}, "\n".join(f"{s['session']}  {s['phase']:8} {s['file']}" for s in items) or "No session yet."
    if args.what == "show":
        return cmd_status(args)
    ws = paths.resolve_workspace(args.ws)
    config = workspace.load(ws)
    if not args.file:
        raise GateError("Say which testcase file: oqa session new --file <path to .xlsx>")
    path = Path(args.file).expanduser().resolve()
    cases = intake.build(path, config["reference_dirs"])
    session = Session.create(ws, path, cases)
    missing = [(c["id"], r["text"]) for c in cases["cases"] for r in c["refs"] if r["status"] == "missing"]
    lines = [f"Session {session.id} in workspace {ws}: {len(cases['cases'])} cases read from {path.name} (sheet {cases['source']['sheet']})."]
    lines += [f"  warning: {w}" for w in cases["warnings"]]
    lines += [f"  reference not found: {case} mentions {name}" for case, name in missing]
    lines += memory.summary(ws, memory.stem_of(path), [c["id"] for c in cases["cases"]])
    lines.append("Next: " + Session.open(ws, session.id).summary()["next"])
    return {"session": session.id, "workspace": ws, "cases": len(cases["cases"]), "warnings": cases["warnings"], "missing_refs": [{"case": c, "file": f} for c, f in missing]}, "\n".join(lines)


def cmd_analyze(args):
    session = Session.open(args.ws, args.session)
    config = workspace.load(session.ws)
    cases = session.cases()
    if args.template:
        session.require("ANALYZE", "CONNECT", "CONFIRM")
        skeleton = analysis.template(cases, config)
        if args.out:
            write_json(args.out, skeleton)
            return {"template": args.out}, f"Wrote the template for {len(skeleton['cases'])} cases to {args.out}. Fill every judgement field, then: oqa analyze --submit {args.out}"
        return skeleton, json.dumps(skeleton, ensure_ascii=False, indent=2)
    if not args.submit:
        raise GateError("Use --template to get the skeleton, or --submit <file> to hand in the analysis.")
    session.require("ANALYZE", "CONNECT", "CONFIRM")
    submitted = read_json(args.submit)
    if not isinstance(submitted, dict):
        raise SchemaError(f"{args.submit} is not an analysis file", "analysis")
    analysis.validate(submitted, cases, config)
    write_json(session.dir / "analysis.json", submitted)
    session.analysis_submitted()
    counts = {}
    for case in submitted["cases"]:
        counts[case["verdict"]] = counts.get(case["verdict"], 0) + 1
    return {"accepted": True, "verdicts": counts}, f"Analysis accepted ({', '.join(f'{n} {v}' for v, n in counts.items())}). Next: {session.summary()['next']}"


def cmd_strategy(args):
    session = Session.open(args.ws, args.session)
    session.require("CONFIRM")
    config = workspace.load(session.ws)
    data, text = analysis.strategy(read_json(session.dir / "analysis.json"), session.cases(), config)
    write_json(session.dir / "strategy.json", data)
    (session.dir / "strategy.md").write_text(text + "\n", encoding="utf-8")
    session.strategy_ready(data["sha"])
    return {"order": data["order"], "excluded": data["excluded"], "file": str(session.dir / "strategy.md"), "sha": data["sha"]}, \
        f"Strategy written to {session.dir / 'strategy.md'}: {len(data['order'])} cases to run, {len(data['excluded'])} held back.\nShow it to the user and ask them to approve it."


def cmd_confirm(args, is_tty=None):
    session = Session.open(args.ws, args.session)
    state = session.require("CONFIRM")
    tty = (sys.stdin.isatty() and sys.stdout.isatty()) if is_tty is None else is_tty
    if not tty:
        raise GateError("Only the user can confirm: run `oqa confirm` yourself in a terminal, or press Confirm in the dashboard. The assistant cannot do this for them.")
    if is_tty is None:
        print((session.dir / "strategy.md").read_text(encoding="utf-8"))
        if input("Type 'yes' to approve this strategy and start testing: ").strip().lower() not in ("yes", "y", "có", "co"):
            raise GateError("Not confirmed.")
    strategy_doc = session.confirm_current("terminal")
    return {"confirmed": True, "by": "terminal"}, f"Confirmed. {len(strategy_doc['order'])} cases will run. Next: {session.summary()['next']}"


def cmd_connect(args):
    session = Session.open(args.ws, args.session)
    ws = workspace.load(session.ws)
    info = connect.connect(session, ws, default_store(), headless=not args.headed)
    lines = [f"Signed in: {name} -> {item['landing_url']}" + (f" (Odoo {item['version']})" if item["version"] else "") for name, item in info.items()]
    return {"profiles": info}, "\n".join(lines) + "\nNext: " + session.summary()["next"]


def cmd_scaffold(args):
    session = Session.open(args.ws, args.session)
    done = scaffold.scaffold(session, args.cases or None)
    advice = {"new": "fill in the actions", "reuse": "same case as before, nothing to do", "changed": "the case changed since this script was written: review it", "unknown": "not in the analysis"}
    return {"scripts": done}, "\n".join(f"{d['case']}: {d['status']} - {advice[d['status']]}" + (f"  {d['path']}" if d["path"] else "")
                                        + (f"\n   WEAK: step(s) {d['unchecked']} have no check; add s.check / s.equals / o.expect_* / assert before running" if d.get("unchecked") else "")
                                        + (f"\n   WEAK: step(s) {d['unhighlighted']} never call s.hl(<selector>); name the element each step is about" if d.get("unhighlighted") else "") for d in done)


def cmd_run(args):
    session = Session.open(args.ws, args.session)
    session.require("EXECUTE")
    ws = workspace.load(session.ws)
    order = read_json(session.dir / "strategy.json")["order"]
    results = session.state()["results"]
    cases = args.cases or [c for c in order if c not in results]
    if not cases:
        return {"results": []}, "Every case already has a result. Next: " + session.summary()["next"]
    if args.inline or os.environ.get("OQA_NO_SERVICE"):
        out = runtime.Runner(session, ws, default_store(), headless=not args.headed).run(cases, start=args.start, fresh=args.fresh)
    else:
        out = service.run(session, ws["name"], cases, args.start, args.fresh, args.headed)
    lines = []
    for r in out:
        lines.append(f"{r['case']} attempt {r['attempt']}: {r['status']} ({r['steps_passed']}/{r['steps_total']} steps)" + (" [script is older than the case]" if r["stale_script"] else ""))
        if r.get("note"):
            lines.append(f"  {r['note']}")
        lines += [f"  user hint: {h}" for h in r["hints"]]
        if r["status"] == "FAIL":
            lines += [f"  failed at step {r['failed_step']}: {r['reason']}"] + ([f"  page says: {r['page_says']}"] if r["page_says"] else []) + [f"  picture: {r['picture']}"] + [f"  known: {k}" for k in r["knowledge"]]
            if r.get("failed_page"):
                lines += [f"  the page where it stopped is saved: oqa probe --failed --case {r['case']} --filter <word>"] + [f"  on that page: {s}" for s in r["similar"]]
    left = out[-1]["left"] if out else []
    lines.append(f"{len(left)} case(s) still without a result" + (": " + ", ".join(left[:8]) if left else ""))
    if out and out[-1]["status"] == "STOPPED":
        lines.append("The user stopped the test. Do nothing until they resume it: run `oqa wait`.")
    elif not out and session.state().get("stopped"):
        lines.append("The user stopped the test. Do nothing until they resume it: run `oqa wait`.")
    elif out and out[-1]["status"] == "FAIL":
        failed = out[-1]
        lines.append(f"Find the cause before touching the script: the picture, `oqa probe --failed --case {failed['case']} --filter <word>` (the page where it stopped), `oqa knowledge search`. "
                     f"Then record it: oqa repair {failed['case']} --cause \"...\" --evidence \"...\" --change \"...\". {failed['runs_left']} run(s) left before you must ask the user.")
        lines.append(f"Fix step {failed['failed_step']} only: `oqa run {failed['case']}` picks up there with steps 1-{max(failed['failed_step'] - 1, 0)} kept, when the script is written as step functions "
                     "(def step_<n>(t, s)). The case's other steps are not run again.")
        if "TimeoutError" in failed["reason"]:
            lines.append("A Timeout means the element did not show up in time: either the selector is wrong (probe it) or the page is only slow (then add --timeout 20 to oqa repair).")
    elif not left:
        lines.append("Next: " + session.summary()["next"])
    return {"results": out}, "\n".join(lines)


def _case_in_strategy(session, case):
    if case not in read_json(session.dir / "strategy.json")["order"]:
        raise GateError(f"{case} is not in the confirmed strategy.")


def _show(found):
    """What `oqa explore` / `oqa try` print about the page: where it is, what it says, and the controls."""
    lines = [f"page: {found['url']}  \"{found['title']}\""]
    if found.get("toasts"):
        lines.append("notices: " + " | ".join(found["toasts"]))
    if found.get("page_says") and "no error message" not in found["page_says"]:
        lines.append("page says: " + found["page_says"])
    lines.append(f"controls ({len(found['controls'])}{'' if found['controls'] else ', none visible'}):")
    return lines + ["  " + c for c in found["controls"]]


def cmd_explore(args):
    session = Session.open(args.ws, args.session)
    session.require("EXECUTE")
    _case_in_strategy(session, args.case)
    found = service.ask(session, workspace.load(session.ws)["name"], "explore", {"case": args.case, "to": args.to, "filter": args.filter or ""}, args.headed)
    if not found["ok"]:
        return found, "\n".join([f"Replay stopped at step {found['step']}: {found['error']}", f"  piece: {found['code']}", "Fix that step again (oqa try --reset) or start it over."] + _show(found))
    head = f"Exploring {args.case}: the page is where step {args.to} begins." if args.to > 1 else f"Exploring {args.case}: a new page at the landing screen. Try step 1 next."
    return found, "\n".join([head] + _show(found) + [f"Next: oqa try {args.case} --step {args.to} --code \"<python for this step>\""])


def cmd_try(args):
    session = Session.open(args.ws, args.session)
    session.require("EXECUTE")
    _case_in_strategy(session, args.case)
    code = args.code
    if args.file:
        code = sys.stdin.read() if args.file == "-" else Path(args.file).read_text(encoding="utf-8")
    if not args.undo and not (code or "").strip():
        raise GateError("Give the code to try: --code \"...\" or --file <path> (or - for standard input).")
    found = service.ask(session, workspace.load(session.ws)["name"], "try", {"case": args.case, "step": args.step, "code": code, "filter": args.filter or "",
                                                                          "reset": args.reset, "undo": args.undo}, args.headed)
    if found.get("undone"):
        return found, f"Dropped the last piece of step {args.step} ({found['pieces']} left). The page did not move back: `oqa explore {args.case} --to {args.step}` replays what is verified."
    if not found["ok"]:
        lines = [f"FAILED on step {args.step} ({found['seconds']}s): {found['error']}", "Nothing was kept and no attempt was used. What the page shows now:"] + _show(found)
        if found["similar"]:
            lines += ["closest to what the code asked for:"] + ["  " + s for s in found["similar"]]
        return found, "\n".join(lines + ["Change the code and try again; `--filter <word>` lists only the matching controls."])
    return found, "\n".join([f"ok: step {args.step} now has {found['pieces']} verified piece(s) ({found['seconds']}s, {found['checks']} check(s) ran)"] + _show(found)
                            + [f"Next: another piece of step {args.step}, or step {args.step + 1}; when every step is tried: oqa build {args.case}"])


def cmd_build(args):
    session = Session.open(args.ws, args.session)
    session.require("EXECUTE", "CONFIRM", "CONNECT", "ANALYZE", working=False)
    item = next((c for c in read_json(session.dir / "analysis.json")["cases"] if c["id"] == args.case), None)
    info = next((c for c in session.cases()["cases"] if c["id"] == args.case), None)
    if item is None or info is None:
        raise GateError(f"{args.case} is not in the analysis.")
    target = scaffold.script_path(session, args.case)
    if target.exists():
        text = target.read_text(encoding="utf-8")
        if not (recipe.sealed(text) or scaffold.unwritten(target) or args.force):
            raise GateError(f"{target.name} has code that was not built from tried steps. Build would replace it: add --force if that is what you want.")
    steps = recipe.steps_of(recipe.load(session, args.case))
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(recipe.seal(recipe.build(session, args.case, item, info["title"], scaffold.header(info, item, scaffold.case_sha(item)))), encoding="utf-8")
    events.append(session.dir, "built", case=args.case, steps=len(steps))
    missing = [s["n"] for s in item["steps"] if s["n"] not in steps]
    weak = scaffold.unchecked_steps(target) + [n for n in scaffold.unhighlighted_steps(target) if n not in scaffold.unchecked_steps(target)]
    lines = [f"Built {target.name}: {len(steps)} of {len(item['steps'])} step(s) from code that ran on the real page."]
    if missing:
        lines.append(f"  not tried yet: step(s) {missing}: oqa try {args.case} --step N --code ...")
    if scaffold.unchecked_steps(target):
        lines.append(f"  no check in step(s) {scaffold.unchecked_steps(target)}: try a piece with s.check(...) / s.equals(...) / o.expect_*")
    if scaffold.unhighlighted_steps(target):
        lines.append(f"  no s.hl in step(s) {scaffold.unhighlighted_steps(target)}: try a piece with s.hl(\"<selector>\")")
    if not (missing or weak):
        lines.append(f"Next: oqa run {args.case}")
    return {"built": str(target), "tried": sorted(steps), "missing": missing}, "\n".join(lines)


def cmd_runner(args):
    session = Session.open(args.ws, args.session)
    if args.what == "serve":                    # started by `oqa run`, in the background; it leaves by itself when the session stops executing or is idle
        ws = workspace.load(session.ws)
        service.RunnerService(session, ws, default_store(), headless=not args.headed).serve(service.announce_to(session))
        return {"stopped": True}, "Runner stopped."
    if args.what == "stop":
        stopped = service.stop(session)
        return {"stopped": stopped}, "Runner stopped: the next oqa run starts a new one." if stopped else "No runner was running."
    info = service.alive(session)
    return {"running": bool(info)}, f"Runner running (pid {info['pid']}, {'busy' if info['busy'] else 'waiting'})" if info else "No runner is running."


def cmd_finish(args):
    session = Session.open(args.ws, args.session)
    session.declare(args.case, args.status, args.reason)
    return {"case": args.case, "status": args.status}, f"{args.case} recorded as {args.status}. Next: " + session.summary()["next"]


def cmd_lookup(args):
    """Read-only questions to the ORM of the Odoo under test: the technical name of a model or field, an external id, a record by name. No script, no step used."""
    from .odoo_ui import Odoo
    try:
        session = Session.open(args.ws, args.session)
        ws = workspace.load(session.ws)
    except paths.PathError:
        session, ws = None, workspace.load(paths.resolve_workspace(args.ws))
    with probe.signed_in(session, ws, default_store(), args.profile, headless=not args.headed) as page:
        o = Odoo(page, ws["url"], ws.get("test_prefix") or "ZZ_AUTOTEST_")
        if args.what == "models":
            data = [{"model": m, "label": n} for m, n in o.models(args.target or "", args.limit)]
            lines = [f"{d['model']}  \"{d['label']}\"" for d in data]
        elif args.what == "fields":
            if not args.target:
                raise GateError("Say the model: oqa lookup fields res.partner [--filter word]")
            data = o.fields(args.target, args.filter or "", args.limit)
            lines = [f"{d['name']}  \"{d.get('string', '')}\"  {d.get('type', '')}" + (f" -> {d['relation']}" if d.get("relation") else "") + ("  required" if d.get("required") else "") + ("  readonly" if d.get("readonly") else "") for d in data]
        elif args.what == "xmlid":
            if not args.target:
                raise GateError("Say the external id: oqa lookup xmlid base.main_company")
            found = o.xmlid(args.target)
            data = [found] if found else []
            lines = [f"{found['ref']} -> {found['model']},{found['res_id']}"] if found else [f"no external id '{args.target}'"]
        elif args.what == "find":
            if not (args.target and args.filter):
                raise GateError("Say the model and the text: oqa lookup find res.partner --filter Azure")
            data = [{"id": i, "name": n} for i, n in o.find(args.target, args.filter, args.limit)]
            lines = [f"{d['id']}  {d['name']}" for d in data]
        else:                                                           # search
            if not args.target:
                raise GateError("Say the model: oqa lookup search sale.order --domain \"[['state','=','sale']]\" --fields name,amount_total")
            import ast as _ast
            domain = _ast.literal_eval(args.domain) if args.domain else []
            data = o.records(args.target, domain, [f for f in (args.fields or "").split(",") if f], args.limit)
            lines = [json.dumps(r, ensure_ascii=False) for r in data]
    return {"results": data}, "\n".join(lines or ["(nothing found)"])


def cmd_probe(args):
    if args.failed:
        session = Session.open(args.ws, args.session)
        try:
            found = probe.failed(session, args.case, args.filter or "")
        except FileNotFoundError as error:
            raise GateError(str(error))
        events.append(session.dir, "probe", url=found["url"], count=len(found["elements"]), **({"filter": args.filter} if args.filter else {}), **({"case": args.case} if args.case else {}))
        return found, f"{found['url']}  \"{found['title']}\"  (the page where the step stopped; {found['total']} controls were saved)\n" + "\n".join(found["elements"])
    try:
        session = Session.open(args.ws, args.session)
        ws = workspace.load(session.ws)
    except paths.PathError:
        session, ws = None, workspace.load(paths.resolve_workspace(args.ws))
    found = probe.probe(session, ws, default_store(), args.profile, args.path, args.filter or "", headless=not args.headed)
    if session:                                 # the log shows the assistant looked at the real page before it changed anything (a probe from a file shows the same)
        events.append(session.dir, "probe", url=found["url"], count=len(found["elements"]), **({"filter": args.filter} if args.filter else {}), **({"case": args.case} if args.case else {}))
    return found, f"{found['url']}  \"{found['title']}\"\n" + "\n".join(found["elements"])


def cmd_repair(args):
    session = Session.open(args.ws, args.session)
    state = session.require("EXECUTE")
    path = scaffold.script_path(session, args.case)
    note = session.repair(args.case, args.cause, args.evidence, args.change, args.reanalysis or "", scaffold.script_sha(path), args.timeout, args.unchanged or "")
    left = max(0, 1 + MAX_REPAIRS - state["attempts"].get(args.case, 0))
    text = f"Repair note recorded for {args.case} (after run {note['attempt']})" + ("" if note["probed"] else ". You did not probe the page since the failure: if the selector was the problem, `oqa probe` first next time")
    return note, text + f". Next: oqa run {args.case}  ({left} run(s) left)"


def cmd_report(args):
    session = Session.open(args.ws, args.session)
    ws = workspace.load(session.ws)
    if args.partial or session.state()["phase"] == "DONE":           # again after the session closed: rebuild it, change nothing
        session.require("EXECUTE", "REPORT", "DONE", working=False)
        result = report.build(session, ws)
    else:
        result = report.finish(session, ws)
    opened = not args.no_open and report.open_file(result["main"])
    result["opened"] = opened
    text = [f"Evidence workbook: {result['main']}" + (" (opened on the user's screen)" if opened else ""),
            ", ".join(f"{k} {v}" for k, v in sorted(result["counts"].items()))]
    return result, "\n".join(text)


def _alive(url):
    try:
        with urllib.request.urlopen(url, timeout=1.5) as response:
            return response.status == 200
    except Exception:
        return False


def _retire(running, wait=4.0):
    """Ask a dashboard that runs older code to leave; one too old to know how is ended by its process id (the address in the pointer answered, so that process is the one)."""
    base, _, token = running["url"].partition("/?t=")
    try:
        urllib.request.urlopen(urllib.request.Request(f"{base}/api/quit?t={token}", data=b"{}", method="POST", headers={"Content-Type": "application/json"}), timeout=3).read()
    except Exception:
        pass
    deadline = time.monotonic() + wait
    while time.monotonic() < deadline and _alive(running["url"]):
        time.sleep(0.2)
    if _alive(running["url"]) and running.get("pid"):
        try:
            os.kill(int(running["pid"]), signal.SIGTERM)
        except OSError:
            pass
        deadline = time.monotonic() + wait
        while time.monotonic() < deadline and _alive(running["url"]):
            time.sleep(0.2)


def cmd_dashboard(args):
    """Serve the dashboard and open it. Only one runs per Home: asking again opens the running one on the session you mean (a new tab), it does not start a second."""
    try:
        select = paths.resolve_session(args.ws, args.session)
    except paths.PathError:
        select = None
    pointer = paths.home() / "dashboard.json"
    running = read_json(pointer, None)
    if running and _alive(running["url"]):
        if running.get("build") == build_id():
            url = running["url"] + (f"#{quote(select[0])}/{quote(select[1])}" if select else "")
            if not args.no_open:
                webbrowser.open(url)
            return {"url": url, "reused": True}, f"Dashboard already running: {url}"
        _retire(running)                         # started from older files: its pages and its answers no longer match, so a new one takes its place

    def announce(url):
        write_json(pointer, {"url": url, "pid": os.getpid(), "build": build_id()})
        print(f"Dashboard: {url}", flush=True)

    try:
        dashboard_server.Dashboard(select, args.timeout).serve(announce, open_browser=not args.no_open, forever=True)
    finally:
        pointer.unlink(missing_ok=True)
    return {"stopped": True}, "Dashboard stopped."


def cmd_hooks(args):
    """The two hooks as a snippet for settings.json, with this skill's real path. The hooks in SKILL.md's frontmatter use ${CLAUDE_PLUGIN_ROOT} (checked on
    Claude Code 2.1.325: it is the skill's folder there; ${CLAUDE_SKILL_DIR} expands to nothing in a hook). Use this where that does not hold."""
    root = Path(__file__).resolve().parents[1]
    command = lambda name: f'python "{root / "hooks" / name}"'
    snippet = {"hooks": {"PreToolUse": [{"matcher": "Bash|Write|Edit|MultiEdit|Read|Grep", "hooks": [{"type": "command", "command": command("guard.py")}]}],
                         "Stop": [{"hooks": [{"type": "command", "command": command("stop.py")}]}]}}
    return snippet, json.dumps(snippet, indent=2)


def cmd_install(args):
    """Copy this skill (without its dev files) into a skills folder, so it can be taken to another machine or to the add-in's skills."""
    root = Path(__file__).resolve().parents[1]
    target = Path(args.to).expanduser() / root.name
    if target.exists():
        if not args.force:
            raise GateError(f"{target} already exists. Use --force to replace it.")
        if not (target / "oqa.py").exists():
            raise GateError(f"{target} does not look like this skill; remove it yourself if you are sure.")
        shutil.rmtree(target)
    shutil.copytree(root, target, ignore=shutil.ignore_patterns(".venv", "tests", "__pycache__", ".pytest_cache", "*.pyc", "pytest.ini"))
    return {"installed": str(target)}, f"Installed to {target}"


def cmd_ask(args):
    session = Session.open(args.ws, args.session)
    session.ask(args.question, args.case)
    return {"waiting": True}, "Question recorded. Show it to the user and stop until they answer; then run `oqa answer`."


def cmd_answer(args):
    session = Session.open(args.ws, args.session)
    asked = session.answer(args.text)
    session.mark_seen()                      # the assistant ran this itself: it is not news to it
    return {"answered": asked["question"]}, "Answer recorded. You can continue."


def _describe(event):
    """One user action, as a sentence that says what to do about it."""
    kind, case = event["kind"], event.get("case")
    if kind == "confirmed":
        return "The user confirmed the strategy. Next: oqa run"
    if kind == "answered":
        return f"The user answered{f' (about {case})' if case else ''}: {event['answer']}"
    if kind == "hint":
        return f"The user sent a hint for {case}: {event['text']}\n{case} is back in the queue with a fresh budget. Next: oqa run {case}"
    if kind == "stopped":
        return "The user stopped the test. Do nothing until they resume it: run `oqa wait` again."
    return "The user resumed the test. Next: oqa run"


def cmd_wait(args):
    """Block until the user does something on the dashboard (confirm, answer, hint, stop, resume), then say what. Run it in the background after handing over to the
    user; the assistant is told when it ends."""
    session = Session.open(args.ws, args.session)
    deadline = time.monotonic() + args.timeout
    while True:
        news = session.news()
        if news:
            session.mark_seen()
            return {"events": news}, "\n".join(_describe(e) for e in news)
        if time.monotonic() > deadline:
            raise GateError(f"The user did nothing in {args.timeout} seconds. Run `oqa wait` again if you are still waiting for them.")
        time.sleep(0.5)


def cmd_stop(args):
    session = Session.open(args.ws, args.session)
    session.stop("terminal")
    return {"stopped": True}, "Stopped. The step in progress finishes; nothing else starts until `oqa resume`."


def cmd_resume(args, is_tty=None):
    session = Session.open(args.ws, args.session)
    if not ((sys.stdin.isatty() and sys.stdout.isatty()) if is_tty is None else is_tty):
        raise GateError("Only the user resumes a test they stopped: run `oqa resume` in your own terminal or press Resume in the dashboard.")
    session.resume("terminal")
    return {"stopped": False}, "Resumed."


def cmd_status(args):
    session = Session.open(args.ws, args.session)
    summary = session.summary()
    lines = [f"{summary['workspace']} / {summary['session']}  phase {summary['phase']}  ({summary['cases_done']}/{summary['cases_planned']} cases done)"]
    lines.append("STOPPED BY THE USER: do nothing until they resume (oqa wait)" if summary["stopped"] else
                 "WAITING FOR THE USER: " + summary["waiting"]["question"] if summary["waiting"] else "Next: " + summary["next"])
    if summary["needs_repair"]:
        lines.append("Failed and not yet explained: " + ", ".join(summary["needs_repair"]) + ". Before running again: oqa repair <case> --cause ... --evidence ... --change ...")
    lines += [_describe(e) for e in session.news()]
    return summary, "\n".join(lines)


def cmd_knowledge(args):
    ws = paths.resolve_workspace(args.ws)
    if args.what == "learn":
        result = knowledge.learn(ws, args.screen, args.kind, args.fact, args.tag, args.ref)
        return {"result": result}, f"Knowledge {result}."
    found = knowledge.search(ws, args.screen, args.kind, args.text)
    return {"items": found}, "\n".join(f"[screen:{i['screen']}] [kind:{i['kind']}] {i['fact']} -> {i['ref']}" for i in found) or "Nothing known yet."


# --- wiring -----------------------------------------------------------------------
def build_parser():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--ws", help="workspace name (default: $OQA_WS, or the only one)")
    common.add_argument("--session", help="session id (default: $OQA_SESSION, or the workspace's current one)")
    common.add_argument("--json", action="store_true", help="print machine-readable JSON")
    parser = argparse.ArgumentParser(prog="oqa", description="odoo-test: test Odoo from a testcase workbook, in one session per request.")
    parser.add_argument("--version", action="version", version=f"oqa {VERSION}")
    sub = parser.add_subparsers(dest="command", required=True)

    def add(name, func, help_, **kw):
        p = sub.add_parser(name, parents=[common], help=help_, **kw)
        p.set_defaults(func=func)
        return p

    p = add("init", cmd_init, "open the configuration form and wait until it is saved")
    p.add_argument("--timeout", type=int, default=900)
    p.add_argument("--no-open", action="store_true", help="do not open the browser, only print the address")
    p.add_argument("--status", action="store_true", help="only say whether a workspace is configured")
    p = add("doctor", cmd_doctor, "check this computer (and --fix what is missing)")
    p.add_argument("--fix", action="store_true")
    p = add("ws", cmd_ws, "show or list workspaces")
    p.add_argument("what", choices=["show", "list"])
    p = add("session", cmd_session, "new / show / list sessions")
    p.add_argument("what", choices=["new", "show", "list"])
    p.add_argument("--file", help="testcase workbook (.xlsx) for `new`")
    p = add("analyze", cmd_analyze, "get the analysis template, or hand in the analysis")
    p.add_argument("--template", action="store_true")
    p.add_argument("--submit", metavar="FILE")
    p.add_argument("--out", metavar="FILE", help="with --template: write it here instead of printing")
    add("strategy", cmd_strategy, "build the strategy the user must approve")
    add("confirm", cmd_confirm, "the USER approves the strategy (needs a real terminal)")
    p = add("connect", cmd_connect, "sign every profile in and record how this Odoo behaves after login")
    p.add_argument("--headed", action="store_true", help="show the browser (for a two-factor code only the user can enter)")
    p = add("scaffold", cmd_scaffold, "write the case script skeletons (reuses scripts from earlier sessions)")
    p.add_argument("cases", nargs="*", help="case ids (default: every case of the strategy)")
    p = add("run", cmd_run, "run cases in one browser; stops at the first failure; a failed case picks up at the step that failed")
    p.add_argument("cases", nargs="*", help="case ids (default: every case without a result, in strategy order)")
    p.add_argument("--headed", action="store_true")
    p.add_argument("--from", dest="start", type=int, metavar="STEP", help="begin the (first) case at this step, on the page the runner kept (default: the step that failed)")
    p.add_argument("--fresh", action="store_true", help="forget what the runner kept: start the case at step 1 on a new page")
    p.add_argument("--inline", action="store_true", help="run in this process and close the browser afterwards (no runner service, nothing is kept for the next run)")
    p = add("explore", cmd_explore, "open the page where a case's step begins, to try its code on the real page before it is written down")
    p.add_argument("case")
    p.add_argument("--to", type=int, default=1, help="the step to be at: the verified steps before it are replayed (default: a new page at step 1)")
    p.add_argument("--filter", help="list only the controls whose description contains this")
    p.add_argument("--headed", action="store_true")
    p = add("try", cmd_try, "run code for a step on the page the exploration is at; kept only when it works")
    p.add_argument("case")
    p.add_argument("--step", type=int, required=True)
    p.add_argument("--code", help="python for the step: t, o (helpers), s (s.hl, s.check, s.equals), page, t.data are in scope")
    p.add_argument("--file", help="read the code from a file (- for standard input)")
    p.add_argument("--filter", help="list only the controls whose description contains this")
    p.add_argument("--reset", action="store_true", help="forget the verified pieces of this step and start it over with this code")
    p.add_argument("--undo", action="store_true", help="drop the last verified piece of this step")
    p.add_argument("--headed", action="store_true")
    p = add("build", cmd_build, "write the case script from the code that was tried on the real page")
    p.add_argument("case")
    p.add_argument("--force", action="store_true", help="replace a script that has hand-written code")
    p = add("runner", cmd_runner, "the background runner that keeps the browser between runs")
    p.add_argument("what", choices=["serve", "stop", "status"])
    p.add_argument("--headed", action="store_true")
    p = add("finish", cmd_finish, "record FAIL (with evidence), BLOCKED or PENDING for a case you stop working on")
    p.add_argument("case")
    p.add_argument("--status", required=True, choices=["FAIL", "BLOCKED", "PENDING"])
    p.add_argument("--reason", required=True)
    p = add("repair", cmd_repair, "after a failed run: record what you found and changed (required before the case runs again)")
    p.add_argument("case")
    p.add_argument("--cause", required=True, help="why the step failed, as found (not a guess)")
    p.add_argument("--evidence", required=True, help="what showed it: the picture, the probe, the page text")
    p.add_argument("--change", required=True, help="what you changed in the script or the data")
    p.add_argument("--reanalysis", help="how you now read the case; required when the same failure came twice")
    p.add_argument("--timeout", type=int, help="seconds one wait may last on this case's pages from now on (the page is slow, not wrong)")
    p.add_argument("--unchanged", help="when the script is unchanged: what else you changed (data created through the UI, a setting)")
    p = add("lookup", cmd_lookup, "ask the Odoo under test (read only): models, fields, external ids, records by name")
    p.add_argument("what", choices=["models", "fields", "xmlid", "find", "search"])
    p.add_argument("target", nargs="?", help="models: text in the name; fields/find/search: the model; xmlid: module.name")
    p.add_argument("--filter", help="fields: text in the field name or label; find: the text to look for")
    p.add_argument("--domain", help="search: a domain, for example \"[['state','=','sale']]\"")
    p.add_argument("--fields", help="search: comma separated fields to return")
    p.add_argument("--limit", type=int, default=20)
    p.add_argument("--profile")
    p.add_argument("--headed", action="store_true")
    p = add("probe", cmd_probe, "list what can be clicked and filled on a page (to find a selector)")
    p.add_argument("--case", help="the case this probe is for (shows in its repair log); with --failed, the case whose failed step to read")
    p.add_argument("--failed", action="store_true", help="read the page where the latest failed step stopped (saved when it failed; no browser is opened)")
    p.add_argument("--profile")
    p.add_argument("--path", help="page to open (default: the landing page)")
    p.add_argument("--filter", help="only elements whose description contains this")
    p.add_argument("--headed", action="store_true")
    p = add("report", cmd_report, "write the Excel evidence per case and the summary; closes the session")
    p.add_argument("--partial", action="store_true", help="write what exists now without closing the session")
    p.add_argument("--no-open", action="store_true", help="do not open the workbook")
    p = add("dashboard", cmd_dashboard, "open the progress dashboard (history, cases, steps, pictures)")
    p.add_argument("--timeout", type=int, default=4 * 3600, help="seconds before it stops by itself")
    p.add_argument("--no-open", action="store_true")
    add("hooks", cmd_hooks, "print the skill's hooks as a settings.json snippet (absolute paths)")
    p = add("install", cmd_install, "copy this skill into a skills folder")
    p.add_argument("--to", required=True, help="the skills folder (e.g. ~/.claude/skills or %%USERPROFILE%%/.excel-partner/skills)")
    p.add_argument("--force", action="store_true", help="replace an existing copy")
    p = add("ask", cmd_ask, "ask the user something and wait")
    p.add_argument("question")
    p.add_argument("--case")
    p = add("wait", cmd_wait, "block until the user acts on the dashboard (confirm, answer, hint, stop, resume), then say what")
    p.add_argument("--timeout", type=int, default=3600)
    add("stop", cmd_stop, "the USER stops the test (the step in progress finishes)")
    add("resume", cmd_resume, "the USER lets a stopped test go on (needs a real terminal)")
    p = add("answer", cmd_answer, "record the user's reply")
    p.add_argument("text")
    add("status", cmd_status, "phase, progress and the next command")
    p = add("knowledge", cmd_knowledge, "search or add what was learned about this Odoo")
    p.add_argument("what", choices=["search", "learn"])
    p.add_argument("--screen")
    p.add_argument("--kind")
    p.add_argument("--text")
    p.add_argument("--fact")
    p.add_argument("--tag", default="")
    p.add_argument("--ref", default="-")
    return parser


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)
    try:
        data, text = args.func(args)
    except REFUSED as error:
        message = "\n".join(error.problems) if isinstance(error, SchemaError) else str(error)
        if args.json:
            print(json.dumps({"refused": True, "error": message}, ensure_ascii=False))
        else:
            print("REFUSED: " + message, file=sys.stderr)
        return 2
    print(json.dumps(data, ensure_ascii=False, indent=2) if args.json else text)
    return 0
