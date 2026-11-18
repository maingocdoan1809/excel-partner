"""Intake, ordering, events, knowledge, workspace and the credential store."""
import sys

import pytest

from conftest import filled, make_workbook
from oqa import analysis, events, intake, knowledge, paths, workspace
from oqa.secrets import MemoryStore, WindowsCredentialStore
from oqa.store import SchemaError, read_json


# --- intake -----------------------------------------------------------------------
def test_intake_reads_steps_references_and_mentioned_cases(workbook, tmp_path):
    (tmp_path / "refs" / "sub").mkdir(parents=True)
    (tmp_path / "refs" / "sub" / "import.xlsx").write_text("x")
    doc = intake.build(workbook, [tmp_path / "refs"])
    by_id = {c["id"]: c for c in doc["cases"]}
    assert doc["source"]["sheet"] == "Cases" and doc["source"]["header_row"] == 2
    assert by_id["ORD-001"]["steps"] == ["Open Sales", "Click New", "Save"]
    assert by_id["ORD-003"]["refs"] == [{"kind": "file", "text": "import.xlsx", "status": "found", "path": str(tmp_path / "refs" / "sub" / "import.xlsx")}]
    assert by_id["ORD-003"]["case_refs"] == ["ORD-001"]


def test_intake_reports_a_missing_file_and_duplicate_ids(tmp_path):
    book = make_workbook(tmp_path / "d.xlsx", [("A-1", "t", "", "1. use data.csv", "ok"), ("A-1", "again", "", "1. x", "ok")])
    doc = intake.build(book)
    assert doc["cases"][0]["refs"][0]["status"] == "missing"
    assert len(doc["cases"]) == 1 and "appears 2 times" in doc["warnings"][0]


def test_intake_understands_japanese_and_vietnamese_headers(tmp_path):
    book = make_workbook(tmp_path / "j.xlsx", [("T-1", "確認", "", "手順1", "期待")], header=("ケース番号/ID", "テスト項目", "事前準備", "手順", "期待結果"))
    assert intake.build(book)["cases"][0]["steps"] == ["手順1"]


def test_intake_refuses_a_file_that_is_not_a_testcase_table(tmp_path):
    book = make_workbook(tmp_path / "n.xlsx", [("a", "b", "c", "d", "e")], header=("x", "y", "z", "w", "v"))
    with pytest.raises(SchemaError, match="no testcase table"):
        intake.build(book)


# --- ordering ---------------------------------------------------------------------
def test_order_respects_dependencies_and_stays_on_one_profile(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config, **{"ORD-001": {"profile": "Admin"}, "ORD-002": {"profile": "Lead"}, "ORD-003": {"profile": "Admin", "depends_on": ["ORD-001"]}})
    ordered, excluded = analysis.order(data, session.cases())
    assert ordered == ["ORD-001", "ORD-003", "ORD-002"] and excluded == []          # ORD-003 jumps ahead of ORD-002: same profile, no re-login


def test_order_holds_back_cases_that_are_not_ready_and_rejects_cycles(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config, **{"ORD-002": {"verdict": "needs_user", "questions": ["How many orders exist?"]}})
    ordered, excluded = analysis.order(data, session.cases())
    assert ordered == ["ORD-001", "ORD-003"] and excluded[0]["id"] == "ORD-002"
    cyclic = filled(session.cases(), config, **{"ORD-001": {"depends_on": ["ORD-002"]}, "ORD-002": {"depends_on": ["ORD-001"]}})
    with pytest.raises(SchemaError, match="cycle"):
        analysis.validate(cyclic, session.cases(), config)


def test_a_ready_case_cannot_depend_on_one_that_is_not_ready(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config, **{"ORD-001": {"verdict": "blocked", "reason": "no data"}, "ORD-003": {"depends_on": ["ORD-001"]}})
    with pytest.raises(SchemaError, match="not ready"):
        analysis.validate(data, session.cases(), config)


# --- events -----------------------------------------------------------------------
def test_a_step_without_a_picture_is_not_recorded(tmp_path):
    step = dict(case="A", attempt=1, n=1, caption="c", expected="e", actual="a")
    for status in ("PASS", "FAIL"):
        with pytest.raises(SchemaError, match="picture"):
            events.append(tmp_path, "step", status=status, images=[], **step)
        with pytest.raises(SchemaError, match="picture"):
            events.append(tmp_path, "step", status=status, images=[{"raw": "r.png"}], **step)
    events.append(tmp_path, "step", status="FAIL", images=[{"raw": "r.png", "annotated": "a.png"}], **step)
    events.append(tmp_path, "step", status="SKIPPED", images=[], **step)
    assert [e["status"] for e in events.read(tmp_path)] == ["FAIL", "SKIPPED"]


def test_events_reject_unknown_kinds_and_missing_fields(tmp_path):
    with pytest.raises(SchemaError):
        events.append(tmp_path, "bogus")
    with pytest.raises(SchemaError, match="needs 'phase'"):
        events.append(tmp_path, "phase", text="x")


# --- knowledge --------------------------------------------------------------------
def test_knowledge_dedupes_and_never_keeps_a_secret(ws):
    assert knowledge.learn("demo", "sale.order", "quirk", "Save enables after a customer is chosen", "save") == "added"
    assert knowledge.learn("demo", "Sale.Order", "quirk", "save enables  after a customer is chosen") == "duplicate"
    with pytest.raises(knowledge.KnowledgeError):
        knowledge.learn("demo", "login", "login", "password: hunter2")
    with pytest.raises(knowledge.KnowledgeError):
        knowledge.learn("demo", "x", "weird", "fact")
    assert [i["screen"] for i in knowledge.search("demo", kind="quirk")] == ["sale.order"]
    assert knowledge.search("demo", text="nothing like this") == []


# --- workspace and secrets ------------------------------------------------------------
def test_workspace_file_never_holds_a_password(ws, store):
    text = (paths.workspace_dir("demo") / "workspace.json").read_text(encoding="utf-8")
    assert "pw-admin" not in text and "pw-lead" not in text
    assert all(p["has_secret"] for p in workspace.public_view(ws, MemoryStore())["profiles"]) is False   # a store that knows nothing says so


def test_a_blank_password_keeps_the_saved_one_and_a_new_profile_needs_one(ws, store):
    store.set("oqa:demo:Admin", "kept")
    store.set("oqa:demo:Lead", "kept2")
    edited = dict(ws, url="http://localhost:9000")
    workspace.save(edited, store)
    assert store.get("oqa:demo:Admin") == "kept"
    with pytest.raises(SchemaError, match="needs a password"):
        workspace.save(dict(ws, profiles=ws["profiles"] + [{"name": "Sales", "login": "s", "note": ""}]), store)
    workspace.save(dict(ws, profiles=[ws["profiles"][0]]), store)
    assert store.get("oqa:demo:Lead") is None                           # the removed profile's secret goes with it


def test_workspace_validation_says_what_to_fix(store):
    with pytest.raises(SchemaError) as refused:
        workspace.save({"name": "bad name", "url": "", "profiles": []}, store)
    text = " ".join(refused.value.problems)
    assert "workspace name" in text and "URL" in text and "Profiles" in text


def test_url_without_a_scheme_gets_http(store):
    saved = workspace.save({"name": "w", "url": "192.0.2.10:8069/", "profiles": [{"name": "A", "login": "a"}]}, store, {"A": "x"})
    assert saved["url"] == "http://192.0.2.10:8069"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows Credential Manager")
def test_windows_credential_store_roundtrip():
    real, ref = WindowsCredentialStore(), "oqa:pytest:roundtrip"
    try:
        real.set(ref, "pässword-密码")
        assert real.get(ref) == "pässword-密码" and real.has(ref)
    finally:
        real.delete(ref)
    assert real.get(ref) is None


def test_corrupt_or_newer_files_are_errors_not_guesses(tmp_path):
    (tmp_path / "x.json").write_text("{not json", encoding="utf-8")
    with pytest.raises(SchemaError, match="not valid JSON"):
        read_json(tmp_path / "x.json")
    (tmp_path / "y.json").write_text('{"schema_version": 99}', encoding="utf-8")
    with pytest.raises(SchemaError, match="newer oqa"):
        read_json(tmp_path / "y.json")
    assert read_json(tmp_path / "missing.json", {}) == {}


def test_names_cannot_escape_the_home_folder():
    for bad in ("..", "../x", "a/b", "", "a b"):
        with pytest.raises(paths.PathError):
            paths.workspace_dir(bad)


def test_oqa_starts_without_pillow_openpyxl_or_playwright():
    """`oqa doctor --fix` is what installs them, so it must be able to start on a bare Python (found on a machine with none of them)."""
    import subprocess
    code = "import sys; [sys.modules.__setitem__(m, None) for m in ('PIL', 'openpyxl', 'playwright')]; from oqa import cli; sys.exit(cli.main(['doctor']))"
    root = __import__("pathlib").Path(__file__).resolve().parents[1]
    result = subprocess.run([sys.executable, "-c", code], cwd=root, capture_output=True, text=True)
    assert "Not ready" in result.stdout and "Traceback" not in result.stderr, result.stderr


def test_unique_names_keep_the_prefix_and_differ_between_runs(monkeypatch):
    from oqa import odoo_ui
    o = odoo_ui.Odoo.__new__(odoo_ui.Odoo)
    o.prefix = "ZZ_AUTOTEST_"
    first = o.unique("draft")
    monkeypatch.setattr(odoo_ui.time, "strftime", lambda fmt: "999999")
    assert first.startswith("ZZ_AUTOTEST_draft-") and o.unique("draft") == "ZZ_AUTOTEST_draft-999999" and o.name("draft") == "ZZ_AUTOTEST_draft"


# --- memory between sessions on the same testcase file ---------------------------------------
def new_session_on(workbook, ws):
    from oqa.state import Session
    return Session.create("demo", workbook, intake.build(workbook, ws["reference_dirs"]))


def execute(session):
    from oqa.store import write_json
    config = workspace.load(session.ws)
    data = filled(session.cases(), config)
    write_json(session.dir / "analysis.json", data)
    session.analysis_submitted()
    session.connected({})
    doc, _ = analysis.strategy(data, session.cases(), config)
    write_json(session.dir / "strategy.json", doc)
    session.strategy_ready(doc["sha"])
    session.confirm_current("t")


def test_the_next_session_on_the_same_file_knows_the_hints_and_how_cases_ended(ws, workbook):
    first = new_session_on(workbook, ws)
    execute(first)
    first.begin_attempt("ORD-001")
    first.finish_case("ORD-001", "FAIL", "Publish needs a placeholder")
    first.hint("ORD-001", "add one placeholder first")
    first.begin_attempt("ORD-002")
    first.finish_case("ORD-002", "PASS")
    second = new_session_on(workbook, ws)
    template = {c["id"]: c for c in analysis.template(second.cases(), workspace.load("demo"))["cases"]}
    assert template["ORD-001"]["previous"]["hints"] == ["add one placeholder first"]
    assert template["ORD-002"]["previous"]["status"] == "PASS" and template["ORD-003"]["previous"] is None
    first.begin_attempt("ORD-001")
    first.finish_case("ORD-001", "BLOCKED", "still no placeholder")
    assert analysis.template(second.cases(), workspace.load("demo"))["cases"][0]["previous"]["status"] == "BLOCKED"     # the latest word wins


def test_hints_from_sessions_that_ran_before_the_memory_existed_are_read_back(ws, workbook):
    from oqa import memory
    old = new_session_on(workbook, ws)
    execute(old)
    old.begin_attempt("ORD-001")
    old.finish_case("ORD-001", "FAIL", "x")
    state = old.state()
    state["hints"] = {"ORD-001": [{"text": "said in an older session", "at": "2026-01-01 00:00:00", "by": "dashboard"}]}
    from oqa.store import write_json
    write_json(old.file, state)
    (memory._file("demo", memory.stem_of(workbook))).unlink(missing_ok=True)
    assert memory.recall("demo", memory.stem_of(workbook), "ORD-001")["hints"] == ["said in an older session"]


def test_session_new_tells_what_the_earlier_runs_of_the_file_already_say(ws, workbook, monkeypatch, capsys):
    from oqa import cli
    first = new_session_on(workbook, ws)
    execute(first)
    first.begin_attempt("ORD-003")
    first.finish_case("ORD-003", "FAIL", "x")
    first.hint("ORD-002", "use the second page")
    assert cli.main(["session", "new", "--file", str(workbook), "--ws", "demo"]) == 0
    out = capsys.readouterr().out
    assert "Did not pass last time: ORD-003 (FAIL)" in out and "hints for: ORD-002" in out


def test_a_run_number_is_never_reused_so_the_report_shows_the_latest_run(ws, workbook, tmp_path):
    from oqa import events, view
    session = new_session_on(workbook, ws)
    execute(session)
    for _ in range(4):
        n = session.begin_attempt("ORD-001")
        events.append(session.dir, "case_started", case="ORD-001", attempt=n)
        events.append(session.dir, "attempt_finished", case="ORD-001", attempt=n, status="FAIL", reason="x")
        if n < 4:                              # a rerun needs the note of why the last run failed
            events.append(session.dir, "repair", case="ORD-001", attempt=n, cause=f"cause {n}", evidence="the picture", change="the selector")
    session.hint("ORD-001", "try the other selector")
    n = session.begin_attempt("ORD-001")
    events.append(session.dir, "case_started", case="ORD-001", attempt=n)
    events.append(session.dir, "attempt_finished", case="ORD-001", attempt=n, status="PASS", reason="")
    session.finish_case("ORD-001", "PASS")
    record = next(c for c in view.case_records(session) if c["id"] == "ORD-001")
    assert [a["n"] for a in record["attempts"]] == [1, 2, 3, 4, 5] and record["attempts"][-1]["status"] == "PASS" and record["status"] == "PASS"
