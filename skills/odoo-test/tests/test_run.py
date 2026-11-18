"""connect -> scaffold -> run -> repair -> finish -> report, against a fake Odoo and a real browser."""
from pathlib import Path

import pytest

pytest.importorskip("playwright")

from conftest import make_workbook  # noqa: E402
from oqa import analysis, connect, evidence, events, intake, odoo_ui, report, resources, runtime, scaffold, view, workspace  # noqa: E402
from oqa.secrets import MemoryStore  # noqa: E402
from oqa.state import GateError, Session  # noqa: E402
from oqa.store import write_json  # noqa: E402
from support import fake_odoo  # noqa: E402

CREATE = """
def run(t):
    o = t.o
    with t.step(1) as s:
        o.menu("Sales", "Orders")
        s.hl(".o_list_view")
        o.expect_visible(".o_list_view")
    with t.step(2) as s:
        o.new()
        s.hl(".o_form_view")
        o.expect_visible(".o_form_view")
    with t.step(3) as s:
        o.fill("name", o.name("order1"))
        o.m2o("partner_id", "Azure")
        s.hl("[name=name]", "[name=partner_id]")
        o.expect_value("name", o.name("order1"))
        o.expect_value("partner_id", "Azure Interior")
    with t.step(4) as s:
        o.save()
        o.expect_text("Saved")
        s.hl(".o_notification")
"""
PAGINATE = """
def run(t):
    o = t.o

    def make(i):
        o.new()
        o.fill("name", o.name(f"bulk{i}"))
        o.save()
        o.page.wait_for_selector(".o_list_view")

    with t.step(1) as s:
        o.reach(lambda: o.pager()[2] > 5, make)
        s.hl(".o_pager")
        s.check(o.pager()[2] > 5, "the list does not hold more than one page")
    with t.step(2) as s:
        o.page.click(".o_pager_next")
        o.expect_text("6-")
        s.hl(".o_pager")
"""
STEPS = {
    "T-1": [("Open Sales > Orders", "The order list shows"), ("Click New", "A blank form opens"), ("Fill the name and the customer", "Both fields hold the values"), ("Save", "A Saved notice appears")],
    "T-2": [("Make more orders than one page holds", "The pager shows more than 5 records"), ("Go to the next page", "The second page shows")],
}


@pytest.fixture(autouse=True)
def quick(monkeypatch):
    monkeypatch.setattr(odoo_ui, "TIMEOUT", 1500)


@pytest.fixture
def fake():
    server, url = fake_odoo.start(orders=3)
    yield url
    server.shutdown()


@pytest.fixture
def live(fake, tmp_path):
    store = MemoryStore()
    ws = workspace.save({"name": "live", "url": fake, "profiles": [{"name": "Admin", "login": "admin"}]}, store, {"Admin": "pw-admin"})
    book = make_workbook(tmp_path / "orders.xlsx", [("T-1", "Create an order", "", "1. a", "ok"), ("T-2", "Paginate", "Needs more than a page", "1. b\n2. c", "ok; continue from T-1")])
    session = Session.create("live", book, intake.build(book))
    return session, ws, store


CAUSES = ["the menu is called Orderz in this Odoo", "the New button is hidden behind a dialog", "a required field was left empty", "the list loads after the pager", "the profile lacks the access right"]


def explain(session, case, n=0):
    """The repair note a rerun needs; the cause differs from note to note and the script is allowed to be the same."""
    return session.repair(case, CAUSES[n % len(CAUSES)], "the picture and the page text", "changed the script", reanalysis="the case needs the form, not the list",
                          script_sha=scaffold.script_sha(scaffold.script_path(session, case)), unchanged_reason="rerun on purpose")


def approve(session, ws, store, strategy_of=None, uses=None):
    """ANALYZE .. CONFIRM with real code paths (connect uses the browser). `uses` is {case: {step: [resource names]}}."""
    data = {"schema_version": 1, "cases": []}
    for case in session.cases()["cases"]:
        steps = STEPS[case["id"]]
        data["cases"].append({"id": case["id"], "profile": "Admin", "verdict": "ready", "clarity": {"score": 4, "notes": ""}, "sufficiency": {"score": 4, "missing": []}, "refs": [],
                              "data_needs": [], "edge_cases": ["too few rows"], "depends_on": ["T-1"] if case["id"] == "T-2" else [], "reuse": "", **(strategy_of or {}).get(case["id"], {}),
                              "steps": [{"n": n, "action": a, "expected": e, "uses": ((uses or {}).get(case["id"]) or {}).get(n, [])} for n, (a, e) in enumerate(steps, 1)]})
    analysis.validate(data, session.cases(), ws)
    resources.sync(session.dir, data, ws["test_prefix"])
    write_json(session.dir / "analysis.json", data)
    session.analysis_submitted()
    connect.connect(session, ws, store)
    doc, text = analysis.strategy(data, session.cases(), ws)
    write_json(session.dir / "strategy.json", doc)
    (session.dir / "strategy.md").write_text(text, encoding="utf-8")
    session.strategy_ready(doc["sha"])
    session.confirm_current("test")
    return data


def write_script(session, case, body, data):
    path = scaffold.script_path(session, case)
    path.parent.mkdir(parents=True, exist_ok=True)
    item = next(c for c in data["cases"] if c["id"] == case)
    path.write_text(f"# oqa-case: {case} sha={scaffold.case_sha(item)}\n{body}", encoding="utf-8")
    return path


def test_connect_signs_in_keeps_the_state_and_learns_the_landing_page(live):
    session, ws, store = live
    write_json(session.dir / "analysis.json", {})
    approve(session, ws, store)
    assert (session.dir / "auth" / "Admin.json").exists()
    assert session.state()["connect"]["version"] == "17.0"
    from oqa import knowledge
    assert "/odoo" in knowledge.search("live", kind="login")[0]["fact"]


def test_connect_with_a_wrong_password_says_what_odoo_said(live):
    session, ws, store = live
    data = {"schema_version": 1, "cases": [{"id": c["id"], "profile": "Admin", "verdict": "ready", "clarity": {"score": 4}, "sufficiency": {"score": 4}, "refs": [], "edge_cases": ["x"],
                                           "steps": [{"n": 1, "action": "a", "expected": "b"}]} for c in session.cases()["cases"]]}
    write_json(session.dir / "analysis.json", data)
    session.analysis_submitted()
    store.set("oqa:live:Admin", "wrong")
    with pytest.raises(GateError, match="Wrong login/password") as refused:
        connect.connect(session, ws, store)
    assert "picture" in str(refused.value) and session.state()["phase"] == "CONNECT"


def test_scaffold_reuses_unchanged_scripts_and_flags_changed_cases(live):
    session, ws, store = live
    data = approve(session, ws, store)
    first = {r["case"]: r["status"] for r in scaffold.scaffold(session)}
    assert first == {"T-1": "new", "T-2": "new"}
    path = scaffold.script_path(session, "T-1")
    assert path.read_text(encoding="utf-8").count("s.todo()") == 4 and "too few rows" in path.read_text(encoding="utf-8")
    assert {r["case"]: r["status"] for r in scaffold.scaffold(session)} == {"T-1": "reuse", "T-2": "reuse"}
    data["cases"][0]["steps"][0]["expected"] = "something else"
    write_json(session.dir / "analysis.json", data)
    assert {r["case"]: r["status"] for r in scaffold.scaffold(session)}["T-1"] == "changed"


def test_a_script_with_unwritten_steps_is_refused_without_using_an_attempt(live):
    session, ws, store = live
    approve(session, ws, store)
    scaffold.scaffold(session)
    with pytest.raises(GateError, match="unwritten"):
        runtime.Runner(session, ws, store).run(["T-1"])
    assert session.state()["attempts"] == {}


def test_cases_run_in_one_browser_and_every_step_has_a_framed_picture(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE, data)
    write_script(session, "T-2", PAGINATE, data)
    results = runtime.Runner(session, ws, store).run(["T-1", "T-2"])
    assert [r["status"] for r in results] == ["PASS", "PASS"], results
    assert session.state()["phase"] == "REPORT" and session.summary()["counts"] == {"PASS": 2}
    steps = [e for e in events.read(session.dir) if e["kind"] == "step"]
    assert len(steps) == 6 and all(e["status"] == "PASS" for e in steps)
    for e in steps:
        pair = e["images"][0]
        assert evidence.check(session.dir / pair["raw"], session.dir / pair["annotated"]) is None


ATTACH = """
def run(t):
    o = t.o
    with t.step(1) as s:
        f = s.fixture("lines.xlsx")
        o.page.evaluate('''() => { document.body.insertAdjacentHTML('beforeend', '<div><input type=file class=o_input_file><span id=picked></span></div>');
            document.querySelector('.o_input_file').onchange = e => { document.getElementById('picked').textContent = e.target.files[0].name; }; }''')
        o.upload(".o_input_file", f)
        s.hl("#picked")
        s.equals(o.page.inner_text("#picked"), "lines.xlsx")
        s.check(f.rows == 3 and f.columns == ["Product", "Qty"], "the file holds the rows the analysis planned")
    with t.step(2) as s:
        s.hl(".o_web_client")
        o.expect_visible(".o_web_client")
    with t.step(3) as s:
        s.hl(".o_web_client")
        o.expect_visible(".o_web_client")
    with t.step(4) as s:
        s.hl(".o_web_client")
        o.expect_visible(".o_web_client")
"""
LINES = {"name": "lines.xlsx", "mode": "guessed", "purpose": "the lines to attach", "spec": {"columns": ["Product", "Qty"], "rows": [["ZZ_AUTOTEST_p-{i}", 2]], "count": 3}}


def test_a_step_attaches_a_prepared_file_and_its_record_names_it(live):
    session, ws, store = live
    data = approve(session, ws, store, {"T-1": {"resources": [LINES]}}, {"T-1": {1: ["lines.xlsx"]}})
    write_script(session, "T-1", ATTACH, data)
    result = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert result["status"] == "PASS", result
    first = next(e for e in events.read(session.dir) if e["kind"] == "step" and e["n"] == 1)
    assert first["resources"] == [{"name": "lines.xlsx", "kind": "xlsx", "mode": "guessed", "sha256": first["resources"][0]["sha256"], "bytes": first["resources"][0]["bytes"], "rows": 3}]
    assert all("resources" not in e for e in events.read(session.dir) if e["kind"] == "step" and e["n"] > 1)
    shown = next(c for c in view.case_records(session) if c["id"] == "T-1")["attempts"][0]["steps"][0]
    assert shown["resources"][0]["name"] == "lines.xlsx" and shown["uses"] == ["lines.xlsx"]


def test_a_prepared_file_that_was_swapped_after_the_approval_fails_the_step_with_a_picture(live):
    session, ws, store = live
    data = approve(session, ws, store, {"T-1": {"resources": [LINES]}}, {"T-1": {1: ["lines.xlsx"]}})
    write_script(session, "T-1", ATTACH, data)
    resources.file_of(session.dir, "T-1", "lines.xlsx").write_bytes(b"not what the user saw")
    result = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert result["status"] == "FAIL" and result["failed_step"] == 1 and "changed" in result["reason"] and result["picture"]


def test_a_failing_step_is_photographed_explained_repaired_and_run_again(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace('o.menu("Sales", "Orders")', 'o.menu("Sales", "Orderz")'), data)
    first = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert first["status"] == "FAIL" and first["failed_step"] == 1 and first["attempt"] == 1 and first["picture"]
    failing = [e for e in events.read(session.dir) if e["kind"] == "step" and e["status"] == "FAIL"]
    assert len(failing) == 1 and failing[0]["images"] and evidence.check(session.dir / failing[0]["images"][0]["raw"], session.dir / failing[0]["images"][0]["annotated"]) is None
    write_script(session, "T-1", CREATE, data)                                    # the repair
    with pytest.raises(GateError, match="no repair note"):                       # a rerun on a guess is refused, and costs no attempt
        runtime.Runner(session, ws, store).run(["T-1"])
    explain(session, "T-1")
    second = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert second["status"] == "PASS" and second["attempt"] == 2
    assert [a["status"] for a in view.case_records(session)[0]["attempts"]] == ["FAIL", "PASS"]


def test_a_step_that_boxes_nothing_cannot_pass(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace('        s.hl(".o_list_view")\n', ""), data)
    with pytest.raises(GateError, match=r"never call s\.hl") as refused:
        runtime.Runner(session, ws, store).run(["T-1"])
    assert "does not use up an attempt" in str(refused.value) and session.state()["attempts"] == {}


def test_a_break_between_steps_and_out_of_order_steps_are_recorded_with_pictures(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace("    with t.step(2) as s:", "    raise RuntimeError('oops')\n    with t.step(2) as s:"), data)
    broke = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert broke["failed_step"] == 2 and "oops" in broke["reason"] and broke["picture"]
    write_script(session, "T-1", CREATE.replace("t.step(1)", "t.step(2)"), data)
    explain(session, "T-1")
    assert "in order" in runtime.Runner(session, ws, store).run(["T-1"])[0]["reason"]


def test_after_the_repair_budget_the_runner_refuses_and_finish_records_the_failure(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace("o.new()", "o.page.click('.nope')"), data)
    for n in range(4):
        if n:
            explain(session, "T-1", n)
        assert runtime.Runner(session, ws, store).run(["T-1"])[0]["status"] == "FAIL"
    explain(session, "T-1", 4)
    with pytest.raises(GateError, match="oqa ask"):
        runtime.Runner(session, ws, store).run(["T-1"])
    session.declare("T-1", "FAIL", "the New button does nothing")
    session.declare("T-2", "BLOCKED", "needs a customer the user must create")
    assert session.state()["phase"] == "REPORT"


def test_fail_cannot_be_declared_without_a_failed_run_and_pass_never_by_hand(live):
    session, ws, store = live
    approve(session, ws, store)
    with pytest.raises(GateError, match="Run it first"):
        session.declare("T-1", "FAIL", "looks wrong")
    with pytest.raises(GateError, match="PASS comes from"):
        session.declare("T-1", "PASS", "trust me")
    with pytest.raises(GateError, match="Say why"):
        session.declare("T-1", "BLOCKED", " ")


def test_the_report_has_an_excel_per_case_with_pictures_and_closes_the_session(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE, data)
    write_script(session, "T-2", PAGINATE, data)
    runtime.Runner(session, ws, store).run(["T-1", "T-2"])
    result = report.finish(session, ws)
    assert session.state()["phase"] == "DONE" and result["counts"] == {"PASS": 2}
    from openpyxl import load_workbook
    book = load_workbook(result["main"])
    assert book.sheetnames == ["Tổng hợp", "T-1", "T-2"] and Path(result["main"]).name == f"Evidence_orders_{session.id}.xlsx"
    case = book["T-1"]
    assert len(case._images) == 4                                                                   # a red-framed picture for every step
    text = " ¦ ".join(str(c.value) for row in case.iter_rows() for c in row if c.value is not None)
    for expected in ("Create an order", "Các bước (file gốc)", "Kết quả mong đợi (file gốc)", "Tình huống biên đã tính", "too few rows", "Đạt: 4 bước", "Open Sales > Orders", "A Saved notice appears"):
        assert expected in text, expected
    summary = book["Tổng hợp"]
    rows = [[c.value for c in row] for row in summary.iter_rows(min_row=9)]
    assert [r[0] for r in rows if r[0]][:3] == ["No.", "T-1", "T-2"] and [r[7] for r in rows if r[0] in ("T-1", "T-2")] == ["PASS", "PASS"]
    assert not (session.dir / "report" / "cases").exists() and [f.name for f in (session.dir / "report").glob("*.xlsx")] == [Path(result["main"]).name]    # one workbook, no per-case files
    assert "T-2 [PASS]" in (session.dir / "report" / "SUMMARY.md").read_text(encoding="utf-8")


def test_evidence_check_catches_a_picture_without_a_red_frame(tmp_path):
    from PIL import Image
    plain = tmp_path / "plain.png"
    Image.new("RGB", (800, 600), (240, 240, 240)).save(plain)
    framed = tmp_path / "framed.png"
    evidence.annotate(plain, [(100, 100, 300, 200)], "Step 1 PASS: x", framed)
    assert evidence.check(plain, framed) is None
    assert "no red frame" in evidence.check(plain, plain)
    assert "missing" in evidence.check(plain, tmp_path / "nope.png")


def test_the_commands_end_to_end_through_the_cli(live, monkeypatch, capsys):
    from oqa import cli
    session, ws, store = live
    monkeypatch.setattr(cli, "default_store", lambda: store)
    data = approve(session, ws, store)
    flags = ["--ws", "live", "--session", session.id]
    assert cli.main(["scaffold", *flags]) == 0 and "T-1: new" in capsys.readouterr().out
    assert cli.main(["run", *flags]) == 2                                          # unwritten steps: refused, no attempt used
    capsys.readouterr()
    write_script(session, "T-1", CREATE.replace('o.menu("Sales", "Orders")', 'o.menu("Sales", "Orderz")'), data)
    write_script(session, "T-2", PAGINATE, data)
    assert cli.main(["run", *flags]) == 0
    out = capsys.readouterr().out
    assert "T-1 attempt 1: FAIL (0/4 steps)" in out and "failed at step 1" in out and "picture:" in out and "3 run(s) left" in out and "oqa repair T-1" in out
    assert cli.main(["probe", "--filter", "Apps", *flags]) == 0 and 'button#apps.dropdown-toggle "Apps"' in capsys.readouterr().out          # the landing page: the apps toggle, nothing else matches
    write_script(session, "T-1", CREATE, data)
    assert cli.main(["run", *flags]) == 2 and "no repair note" in capsys.readouterr().err          # not before the cause is written down
    assert cli.main(["repair", "T-1", "--cause", "the menu is called Orderz, the real one is Orders", "--evidence", "the probe lists Orders under Sales", "--change", "menu(Sales, Orders)", *flags]) == 0
    note = capsys.readouterr().out
    assert "Next: oqa run T-1" in note and "You did not probe" not in note                     # the probe above came after the failure, and the note says so
    assert cli.main(["run", *flags]) == 0
    out = capsys.readouterr().out
    assert "T-1 attempt 2: PASS" in out and "T-2 attempt 1: PASS" in out and "Next: oqa report" in out
    assert cli.main(["report", "--partial", *flags]) == 0 and session.state()["phase"] == "REPORT"       # partial writes the files without closing
    assert cli.main(["report", *flags]) == 0 and session.state()["phase"] == "DONE"
    assert cli.main(["report", *flags]) == 0 and session.state()["phase"] == "DONE"            # asked again after the session closed: rebuilt, nothing else changes


STOP_MIDWAY = CREATE.replace('        o.new()\n        s.hl(".o_form_view")', '        o.new()\n        s.hl(".o_form_view")\n        t.session.stop("test")')


def test_the_users_stop_finishes_the_step_starts_no_other_and_costs_no_attempt(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", STOP_MIDWAY, data)
    result = runtime.Runner(session, ws, store).run(["T-1", "T-2"])
    assert [r["status"] for r in result] == ["STOPPED"] and result[0]["steps_passed"] == 2          # step 2 was finished, step 3 never began
    assert session.state()["attempts"] == {} and "T-1" not in session.state()["results"]
    steps = [e for e in events.read(session.dir) if e["kind"] == "step"]
    assert [e["status"] for e in steps] == ["PASS", "PASS"]                                         # no FAIL step was invented
    assert view.case_records(session)[0]["status"] == "STOPPED"
    session.resume("test")
    write_script(session, "T-1", CREATE, data)
    again = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert again["status"] == "PASS" and again["attempt"] == 1
    assert view.case_records(session)[0]["attempts"][0]["status"] == "PASS"                         # the later run replaces the stopped one in the view


def test_a_hint_reaches_the_assistant_in_the_next_run_of_the_case(live, monkeypatch, capsys):
    from oqa import cli
    session, ws, store = live
    monkeypatch.setattr(cli, "default_store", lambda: store)
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace("o.new()", "o.page.click('.nope')"), data)
    for n in range(4):
        if n:
            explain(session, "T-1", n)
        assert runtime.Runner(session, ws, store).run(["T-1"])[0]["status"] == "FAIL"
    session.hint("T-1", "the New button is called Create in this list")
    write_script(session, "T-1", CREATE.replace("o.new()", "o.page.click('.nope')"), data)
    assert cli.main(["run", "T-1", "--ws", "live", "--session", session.id]) == 0
    out = capsys.readouterr().out
    assert "user hint: the New button is called Create in this list" in out and "attempt 5: FAIL" in out and "3 run(s) left" in out


def run_t1(live, body, case="T-1"):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, case, body, data)
    return session, runtime.Runner(session, ws, store).run([case])[0]


def test_a_script_with_an_unchecked_step_is_refused_without_using_an_attempt(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace('        o.expect_visible(".o_list_view")\n', ""), data)
    assert scaffold.unchecked_steps(scaffold.script_path(session, "T-1")) == [1]
    with pytest.raises(GateError, match=r"step\(s\) \[1\] have no check"):
        runtime.Runner(session, ws, store).run(["T-1"])
    assert session.state()["attempts"] == {}
    flagged = {r["case"]: r.get("unchecked") for r in scaffold.scaffold(session)}
    assert flagged["T-1"] == [1]                                                     # the reused script is flagged at scaffold time too


def test_a_check_that_never_ran_does_not_count(live):
    body = CREATE.replace('        o.expect_visible(".o_list_view")\n', '        if o.page.url == "never":\n            s.check(True)\n')
    session, result = run_t1(live, body)
    assert result["status"] == "FAIL" and result["failed_step"] == 1 and "no check ran" in result["reason"]


def test_asserts_and_checks_are_counted_and_report_what_they_found(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE.replace('        o.expect_visible(".o_list_view")\n', '        assert o.rows() > 99, f"only {o.rows()} rows"\n'), data)
    result = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert result["status"] == "FAIL" and "only 3 rows" in result["reason"]
    write_script(session, "T-1", CREATE.replace('        o.expect_visible(".o_list_view")\n', '        s.equals(o.rows(), 99, "row count")\n'), data)
    explain(session, "T-1")
    again = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert "row count is 3, expected 99" in again["reason"]


def test_every_passed_step_records_how_many_checks_ran(live):
    session, result = run_t1(live, CREATE.replace('        o.expect_visible(".o_list_view")\n', '        assert o.rows() >= 0\n        s.check(True)\n'))
    assert result["status"] == "PASS", result
    checks = {e["n"]: e["checks"] for e in events.read(session.dir) if e["kind"] == "step"}
    assert checks[1] == 2 and checks[2] == 1 and checks[3] == 2 and checks[4] >= 1       # asserts, s.check and o.expect_* all count
    assert all(s["checks"] for s in view.case_records(session)[0]["attempts"][0]["steps"])


def test_the_scaffold_asks_for_a_check_in_every_step(live):
    session, ws, store = live
    approve(session, ws, store)
    scaffold.scaffold(session)
    text = scaffold.script_path(session, "T-1").read_text(encoding="utf-8")
    assert text.count("every step needs at least one check") == 4


def test_rebuilding_the_report_while_excel_has_it_open_does_not_fail(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", CREATE, data)
    write_script(session, "T-2", PAGINATE, data)
    runtime.Runner(session, ws, store).run(["T-1", "T-2"])
    first = report.build(session, ws)
    folder = Path(first["folder"])
    (folder / ("~$" + Path(first["main"]).name)).write_text("excel lock", encoding="utf-8")      # Excel's lock file sits beside an open workbook
    stale = folder / "T-1.xlsx"
    stale.write_bytes(b"old per-case file")
    with open(stale, "rb"):                                                                          # on Windows a file in use cannot be deleted
        again = report.build(session, ws)                                                            # must neither crash nor touch the lock file
    assert Path(again["main"]).exists() and (folder / ("~$" + Path(first["main"]).name)).exists()
