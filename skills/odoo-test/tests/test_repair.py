"""A failed run must be explained before it is run again: the repair note, what it refuses, what it records and where the user sees it. No browser here."""
import time

import pytest

from test_flow import executing
from oqa import analysis, cli, events, memory, runtime, scaffold, view, workspace
from oqa.secrets import MemoryStore
from oqa.state import GateError, Session
from oqa.store import SchemaError

CASE = "ORD-001"
GOOD = dict(cause="Save is inside the gear menu in Odoo 19, not a toolbar button", evidence="probe on the form lists no button named Save, the gear has it",
            change="open the gear menu then click Save")


def failed_run(session, case=CASE, reason="TimeoutError: waiting for locator .o_form_button_save", sha="aaa"):
    attempt = session.begin_attempt(case)
    events.append(session.dir, "attempt_finished", case=case, attempt=attempt, status="FAIL", reason=reason, failed_step=1, script_sha=sha)
    return attempt


def repair(session, case=CASE, **kw):
    return session.repair(case, **{**GOOD, "script_sha": "bbb", **kw})


@pytest.fixture
def running(session, tmp_path):
    executing(session, tmp_path)
    return session


def test_a_failed_case_cannot_run_again_until_the_repair_is_explained(running):
    failed_run(running)
    with pytest.raises(GateError, match="no repair note was recorded") as refused:
        running.begin_attempt(CASE)
    assert "oqa repair ORD-001" in str(refused.value)
    assert running.summary()["needs_repair"] == [CASE]
    repair(running)
    assert running.summary()["needs_repair"] == []
    assert running.begin_attempt(CASE) == 2


def test_a_case_that_never_failed_has_nothing_to_repair(running):
    with pytest.raises(GateError, match="no failed run"):
        repair(running)
    running.begin_attempt(CASE)
    events.append(running.dir, "attempt_finished", case=CASE, attempt=1, status="PASS", reason="")
    with pytest.raises(GateError, match="no failed run"):
        repair(running)


def test_the_note_must_say_something_and_is_taken_once_per_failure(running):
    failed_run(running)
    with pytest.raises(GateError, match="--evidence says too little"):
        repair(running, evidence="looked")
    repair(running)
    with pytest.raises(GateError, match="already recorded"):
        repair(running, cause="a different cause entirely, found later")


def test_a_cause_that_repeats_an_earlier_one_is_refused(running):
    failed_run(running)
    repair(running)
    running.begin_attempt(CASE)
    events.append(running.dir, "attempt_finished", case=CASE, attempt=2, status="FAIL", reason="the form did not save", script_sha="bbb")
    with pytest.raises(GateError, match="the cause is something else"):
        repair(running, cause="Save is inside the gear menu in Odoo 19, not in a toolbar button", script_sha="ccc")
    repair(running, cause="the Name field is required and was left empty, so Save is refused", script_sha="ccc")


def test_a_repair_that_changed_nothing_is_refused_unless_it_says_what_else_changed(running):
    failed_run(running, sha="same")
    with pytest.raises(GateError, match="same as the one that just failed"):
        repair(running, script_sha="same")
    note = repair(running, script_sha="same", unchanged_reason="created the customer through the UI first")
    assert note["unchanged_reason"] == "created the customer through the UI first"


def test_a_slow_page_gets_a_longer_wait_for_that_case_only(running):
    failed_run(running, sha="same")
    with pytest.raises(GateError, match="seconds from 5 to 120"):
        repair(running, script_sha="same", timeout_s=500)
    repair(running, script_sha="same", timeout_s=20)
    assert running.state()["timeouts"] == {CASE: 20}


def test_the_same_failure_twice_needs_the_case_read_again(running):
    reason = "TimeoutError: waiting for locator .o_data_row after 8000ms"
    failed_run(running, reason=reason)
    repair(running)
    running.begin_attempt(CASE)
    events.append(running.dir, "attempt_finished", case=CASE, attempt=2, status="FAIL", reason=reason.replace("8000", "9000"), script_sha="bbb")
    other = dict(cause="the list opens in kanban and has no table rows", script_sha="ccc")
    with pytest.raises(GateError, match="same failure came twice"):
        repair(running, **other)
    note = repair(running, reanalysis="the case assumes a list view but Templates opens in kanban: switch view first", **other)
    assert note["reanalysis"].startswith("the case assumes")


def test_the_note_records_whether_the_page_was_looked_at_since_the_failure(running):
    failed_run(running)
    assert repair(running)["probed"] is False
    running.begin_attempt(CASE)
    events.append(running.dir, "attempt_finished", case=CASE, attempt=2, status="FAIL", reason="x is missing", script_sha="bbb")
    events.append(running.dir, "probe", url="http://x/odoo", count=12, filter="Save")
    assert repair(running, cause="the Name field is required and was left empty", script_sha="ccc")["probed"] is True


def test_what_was_tried_is_remembered_for_the_next_session_on_the_file(running):
    failed_run(running)
    repair(running)
    stem = memory.stem_of(running.state()["file"])
    tried = memory.recall(running.ws, stem, CASE)["tried"]
    assert tried == [f"{GOOD['cause']} -> {GOOD['change']}"]
    config = workspace.load(running.ws)
    item = next(c for c in analysis.template(running.cases(), config)["cases"] if c["id"] == CASE)
    assert item["previous"]["tried"] == tried


def test_the_repair_log_reaches_the_dashboard_view(running):
    failed_run(running)
    repair(running, timeout_s=15)
    record = next(r for r in view.case_records(running) if r["id"] == CASE)
    assert [(r["attempt"], r["cause"], r["timeout_s"]) for r in record["repairs"]] == [(1, GOOD["cause"], 15)] and record["timeout_s"] == 15
    assert record["status"] == "REPAIRING"


def test_a_session_nobody_touched_for_hours_shows_its_running_cases_as_interrupted(running):
    failed_run(running)
    state, log = running.state(), events.read(running.dir)
    assert view.idle_hours(state, log) < 1
    later = time.time() + 7 * 3600
    hours = view.idle_hours(state, log, now=later)
    assert hours >= view.IDLE_HOURS
    records = view.interrupted([{"status": "REPAIRING"}, {"status": "PASS"}, {"status": "RUNNING"}], hours)
    assert [r["status"] for r in records] == ["INTERRUPTED", "PASS", "INTERRUPTED"]
    assert view.interrupted([{"status": "RUNNING"}], 1)[0]["status"] == "RUNNING"
    running.state()["phase"] = "DONE"
    assert view.idle_hours({"phase": "DONE", "created": "2020-01-01 00:00:00"}, [], now=later) == 0          # a finished session is done, not idle


def test_timeout_s_in_the_analysis_is_a_whole_number_of_seconds_in_range(session, ws):
    data = analysis.template(session.cases(), ws)
    for case in data["cases"]:
        case.update(profile="Admin", verdict="ready", clarity={"score": 4, "notes": ""}, sufficiency={"score": 4, "missing": []}, edge_cases=["x"], refs=[])
        for step in case["steps"]:
            step["expected"] = "ok"
    data["cases"][0]["timeout_s"] = 30
    analysis.validate(data, session.cases(), ws)
    for bad in (2, 500, "20", 8.5, True):
        data["cases"][0]["timeout_s"] = bad
        with pytest.raises(SchemaError, match="timeout_s"):
            analysis.validate(data, session.cases(), ws)


STEPS_OK = '''
def run(t):
    with t.step(1) as s:
        s.hl(".o_form_view")
        s.check(True)
    with t.step(2) as s:
        if t.o.rows():
            s.hl(".o_list_view")
        s.check(True)
'''


def test_a_step_with_no_hl_is_found_by_reading_the_script(tmp_path):
    path = tmp_path / "case.py"
    path.write_text(STEPS_OK, encoding="utf-8")
    assert scaffold.unhighlighted_steps(path) == []
    path.write_text(STEPS_OK.replace('        s.hl(".o_form_view")\n', ""), encoding="utf-8")
    assert scaffold.unhighlighted_steps(path) == [1]
    path.write_text("def run(t:\n", encoding="utf-8")
    assert scaffold.unhighlighted_steps(path) == []                           # a syntax error is the runner's to report, with its line


def test_the_runner_refuses_a_script_with_an_unhighlighted_step_and_uses_no_attempt(running, ws):
    path = scaffold.script_path(running, CASE)
    path.parent.mkdir(parents=True, exist_ok=True)
    steps = len(next(c for c in analysis.template(running.cases(), ws)["cases"] if c["id"] == CASE)["steps"])
    body = "def run(t):\n" + "".join(f"    with t.step({n}) as s:\n        s.check(True)\n" for n in range(1, steps + 1))
    path.write_text("# oqa-case: ORD-001 sha=x\n" + body, encoding="utf-8")
    with pytest.raises(GateError, match=r"never call s\.hl") as refused:
        runtime.Runner(running, ws, MemoryStore()).run_case(None, CASE)
    assert "does not use up an attempt" in str(refused.value) and running.state()["attempts"] == {}


def test_the_repair_command_end_to_end(running, capsys):
    flags = ["--ws", "demo", "--session", running.id]
    failed_run(running)
    assert cli.main(["status", *flags]) == 0 and "Failed and not yet explained: ORD-001" in capsys.readouterr().out
    assert cli.main(["repair", CASE, "--cause", "short", "--evidence", "x", "--change", "y", *flags]) == 2
    assert "--cause says too little" in capsys.readouterr().err
    scaffold.script_path(running, CASE).parent.mkdir(parents=True, exist_ok=True)
    scaffold.script_path(running, CASE).write_text("def run(t):\n    pass\n", encoding="utf-8")
    assert cli.main(["repair", CASE, "--cause", GOOD["cause"], "--evidence", GOOD["evidence"], "--change", GOOD["change"], *flags]) == 0
    out = capsys.readouterr().out
    assert "Repair note recorded for ORD-001 (after run 1)" in out and "You did not probe" in out and "Next: oqa run ORD-001" in out


def test_the_excel_report_tells_how_the_assistant_re_analysed_and_repaired(running, ws):
    from openpyxl import load_workbook

    from oqa import report
    failed_run(running)
    repair(running)
    book = load_workbook(report.build(running, ws)["main"])
    text = " ¦ ".join(str(c.value) for row in book[CASE].iter_rows() for c in row if c.value is not None)
    assert GOOD["cause"] in text and GOOD["evidence"] in text and GOOD["change"] in text and "AI phân tích lại và sửa" in text
