"""The state machine end to end: every command refused out of turn, the user's confirmation, the repair budget and the question gate."""
import pytest

from conftest import filled, write
from oqa import analysis, cli, events, workspace
from oqa.state import MAX_REPAIRS, GateError, Session
from oqa.store import SchemaError, read_json, write_json


def to_confirm(session, tmp_path, **per_case):
    config = workspace.load(session.ws)
    data = filled(session.cases(), config, **per_case)
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", "demo", "--session", session.id]) == 0
    session.connected({"profiles": ["Admin"]})
    assert cli.main(["strategy", "--ws", "demo", "--session", session.id]) == 0
    return data


def confirm(session):
    return cli.cmd_confirm(type("A", (), {"ws": session.ws, "session": session.id})(), is_tty=True)


def test_new_session_starts_in_analyze_and_everything_later_is_refused(session):
    assert session.state()["phase"] == "ANALYZE"
    for gate in (lambda: session.connected({}), lambda: session.strategy_ready("x"), lambda: session.confirm("t", "x"), lambda: session.begin_attempt("ORD-001"),
                 lambda: session.finish_case("ORD-001", "PASS")):
        with pytest.raises(GateError):
            gate()
    assert cli.main(["strategy", "--ws", "demo", "--session", session.id]) == 2


def test_an_incomplete_analysis_is_refused_with_every_defect(session, tmp_path):
    config = workspace.load("demo")
    data = analysis.template(session.cases(), config)           # nothing judged yet
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, session.cases(), config)
    text = "\n".join(refused.value.problems)
    assert "verdict" in text and "edge_cases" in text and "clarity.score" in text and "profile" in text
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", "demo", "--session", session.id]) == 2
    assert session.state()["phase"] == "ANALYZE"


def test_a_ready_case_with_a_missing_reference_or_low_score_is_refused(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config)
    data["cases"][2]["refs"] = [{"text": "import.xlsx", "status": "missing", "note": ""}]
    data["cases"][1]["clarity"]["score"] = 2
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, session.cases(), config)
    assert any("ORD-003" in p and "reference" in p for p in refused.value.problems)
    assert any("ORD-002" in p and "clarity" in p for p in refused.value.problems)


def test_full_flow_and_only_a_person_at_a_terminal_can_confirm(session, tmp_path):
    to_confirm(session, tmp_path)
    assert session.state()["phase"] == "CONFIRM"
    assert cli.main(["confirm", "--ws", "demo", "--session", session.id]) == 2        # pytest has no terminal: the assistant cannot confirm
    assert session.state()["phase"] == "CONFIRM"
    confirm(session)
    assert session.state()["phase"] == "EXECUTE" and session.state()["confirmation"]["by"] == "terminal"
    assert any(e["kind"] == "confirmed" for e in events.read(session.dir))


def test_a_tampered_strategy_cannot_be_confirmed(session, tmp_path):
    to_confirm(session, tmp_path)
    doc = read_json(session.dir / "strategy.json")
    doc["order"] = list(reversed(doc["order"]))
    write_json(session.dir / "strategy.json", doc)
    with pytest.raises(GateError):
        confirm(session)
    assert session.state()["phase"] == "CONFIRM"


def test_changing_the_analysis_after_the_strategy_forces_a_new_strategy(session, tmp_path):
    data = to_confirm(session, tmp_path)
    data["cases"][0]["edge_cases"] = ["a new edge case"]
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "b.json", data)), "--ws", "demo", "--session", session.id]) == 0
    assert session.state()["phase"] == "CONFIRM" and session.state()["strategy_sha"] is None
    with pytest.raises(GateError):
        confirm(session)


def test_repair_budget_then_the_assistant_must_ask(session, tmp_path):
    to_confirm(session, tmp_path)
    confirm(session)
    assert [session.begin_attempt("ORD-001") for _ in range(1 + MAX_REPAIRS)] == [1, 2, 3, 4]
    with pytest.raises(GateError, match="oqa ask"):
        session.begin_attempt("ORD-001")
    session.ask("Which customer should the order use?", "ORD-001")
    with pytest.raises(GateError, match="waiting"):
        session.begin_attempt("ORD-002")
    session.answer("Use Azure Interior")
    assert session.state()["attempts"].get("ORD-001") is None          # fresh budget after the user's answer
    assert session.begin_attempt("ORD-001") == 5                      # ...but the run NUMBER goes on: run 1 of the new budget is run 5 of the case


def test_a_case_outside_the_strategy_or_already_passed_cannot_run(session, tmp_path):
    to_confirm(session, tmp_path)
    confirm(session)
    with pytest.raises(GateError, match="not in the confirmed strategy"):
        session.begin_attempt("NOPE-1")
    session.begin_attempt("ORD-001")
    session.finish_case("ORD-001", "PASS")
    with pytest.raises(GateError, match="already passed"):
        session.begin_attempt("ORD-001")


def test_finishing_every_case_moves_to_report(session, tmp_path):
    to_confirm(session, tmp_path)
    confirm(session)
    for case in session.summary() and read_json(session.dir / "strategy.json")["order"]:
        session.begin_attempt(case)
        session.finish_case(case, "PASS")
    assert session.state()["phase"] == "REPORT"
    assert session.summary()["counts"] == {"PASS": 3}


def test_ask_blocks_the_work_commands_until_answered(session, tmp_path):
    assert cli.main(["ask", "Which Odoo version?", "--ws", "demo", "--session", session.id]) == 0
    assert cli.main(["status", "--ws", "demo", "--session", session.id]) == 0          # reading is always allowed
    data = filled(session.cases(), workspace.load("demo"))
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", "demo", "--session", session.id]) == 2
    assert cli.main(["answer", "17.0", "--ws", "demo", "--session", session.id]) == 0
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", "demo", "--session", session.id]) == 0


def test_session_needs_a_configured_workspace(tmp_path, workbook):
    assert cli.main(["session", "new", "--file", str(workbook)]) == 2


def test_session_commands_use_the_current_session_without_flags(session, capsys):
    assert cli.main(["status", "--json"]) == 0
    assert session.id in capsys.readouterr().out


def test_install_copies_the_skill_without_its_dev_files(tmp_path, capsys):
    target = tmp_path / "skills"
    assert cli.main(["install", "--to", str(target)]) == 0
    copied = target / "odoo-test"
    assert (copied / "SKILL.md").exists() and (copied / "oqa.py").exists() and (copied / "hooks" / "guard.py").exists() and (copied / "oqa" / "web" / "dashboard.html").exists()
    assert not (copied / "tests").exists() and not (copied / ".venv").exists()
    assert cli.main(["install", "--to", str(target)]) == 2                       # will not overwrite without --force
    assert cli.main(["install", "--to", str(target), "--force"]) == 0


def executing(session, tmp_path):
    config = workspace.load(session.ws)
    data = filled(session.cases(), config)
    write_json(session.dir / "analysis.json", data)
    session.analysis_submitted()
    session.connected({})
    doc, _ = analysis.strategy(data, session.cases(), config)
    write_json(session.dir / "strategy.json", doc)
    session.strategy_ready(doc["sha"])
    session.confirm_current("t")
    return doc["order"]


def test_stop_blocks_every_working_command_until_the_user_resumes(session, tmp_path, capsys):
    executing(session, tmp_path)
    session.stop("dashboard")
    with pytest.raises(GateError, match="stopped"):
        session.begin_attempt("ORD-001")
    with pytest.raises(GateError, match="stopped"):
        session.require("EXECUTE")
    with pytest.raises(GateError, match="Already"):
        session.stop("dashboard")
    assert cli.main(["run", "--ws", "demo", "--session", session.id]) == 2
    assert cli.main(["resume", "--ws", "demo", "--session", session.id]) == 2        # the assistant has no terminal: it cannot resume
    assert session.state()["stopped"]
    session.resume("dashboard")
    assert session.begin_attempt("ORD-001") == 1
    with pytest.raises(GateError, match="not stopped"):
        session.resume("dashboard")


def test_a_hint_puts_a_failed_case_back_in_the_queue_with_a_fresh_budget(session, tmp_path):
    executing(session, tmp_path)
    for _ in range(4):
        session.begin_attempt("ORD-001")
    with pytest.raises(GateError, match="oqa ask"):
        session.begin_attempt("ORD-001")
    session.finish_case("ORD-001", "FAIL", "gave up")
    session.hint("ORD-001", "  the Save button is in the Actions menu  ")
    state = session.state()
    assert "ORD-001" not in state["results"] and "ORD-001" not in state["attempts"]
    assert state["hints"]["ORD-001"][0]["text"] == "the Save button is in the Actions menu"
    assert session.begin_attempt("ORD-001") == 5 and session.state()["attempts"]["ORD-001"] == 1


def test_a_hint_after_the_last_case_reopens_execution_and_answers_a_pending_question(session, tmp_path):
    order = executing(session, tmp_path)
    for case in order:
        session.begin_attempt(case)
        session.finish_case(case, "PASS" if case == "ORD-001" else "BLOCKED", "x")
    assert session.state()["phase"] == "REPORT"
    session.hint("ORD-002", "use the Azure Interior customer")
    assert session.state()["phase"] == "EXECUTE"
    session.ask("Which customer for ORD-003?", "ORD-003")
    session.hint("ORD-003", "Azure Interior")
    assert session.state()["waiting"] is None


def test_a_hint_must_be_for_a_planned_case_that_has_not_passed(session, tmp_path):
    with pytest.raises(GateError, match="already ran"):
        session.hint("ORD-001", "x")                                               # still analysing
    executing(session, tmp_path)
    with pytest.raises(GateError, match="not in the confirmed strategy"):
        session.hint("NOPE", "x")
    with pytest.raises(GateError, match="Write the hint"):
        session.hint("ORD-001", "   ")
    session.begin_attempt("ORD-001")
    session.finish_case("ORD-001", "PASS")
    with pytest.raises(GateError, match="already passed"):
        session.hint("ORD-001", "x")


def test_wait_reports_what_the_user_did_once_and_times_out_when_nothing_happens(session, tmp_path, capsys):
    executing(session, tmp_path)
    flags = ["--ws", "demo", "--session", session.id]
    assert cli.main(["wait", "--timeout", "1", *flags]) == 0                       # the confirmation done above is the first news
    assert "confirmed the strategy" in capsys.readouterr().out
    assert cli.main(["wait", "--timeout", "1", *flags]) == 2                       # nothing new
    session.begin_attempt("ORD-001")
    session.hint("ORD-001", "try the Actions menu")
    session.stop("dashboard")
    capsys.readouterr()
    assert cli.main(["wait", "--timeout", "1", *flags]) == 0
    out = capsys.readouterr().out
    assert "hint for ORD-001: try the Actions menu" in out and "back in the queue" in out and "stopped the test" in out
    assert cli.main(["status", *flags]) == 0 and "STOPPED BY THE USER" in capsys.readouterr().out
