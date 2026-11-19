"""A case is not ready on a guess: vague wording must be made concrete from a file or an answer, files are shared by path (never uploaded), a case that needs the user
holds the session in ANALYZE with its questions asked, and blocked / needs_user say what was tried first."""
import openpyxl
import pytest
from openpyxl.comments import Comment

from conftest import filled, make_workbook, write
from oqa import analysis, cli, dashboard_server, events, intake, refs, workspace
from oqa.state import Session
from oqa.store import SchemaError


@pytest.fixture
def design(tmp_path):
    folder = tmp_path / "Thu muc thiet ke"
    folder.mkdir()
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Man hinh"
    sheet.append(["Truong", "Nhan", "Thu tu"])
    sheet.append(["partner_id", "Khach hang", 1])
    sheet.append(["date_order", "Ngay dat", 2])
    sheet["B2"].comment = Comment("hien thi truoc ngay dat", "x")
    path = folder / "BD man hinh.xlsx"
    book.save(path)
    return path


@pytest.fixture
def vague_session(ws, tmp_path):
    book = make_workbook(tmp_path / "v.xlsx", [
        ("V-1", "Screen", "", "1. Open the order form", "Màn hình hiển thị đúng design"),
        ("V-2", "Works", "", "1. Click Save", "The order works correctly"),
        ("V-3", "Plain", "", "1. Click Save", "Order is saved and the status shows Draft"),
        ("V-4", "Empty", "", "1. Click Save", ""),
        ("V-5", "Attach", "", "1. Upload the attachment", "File is listed"),
    ])
    return Session.create("demo", book, intake.build(book, ws["reference_dirs"]))


def flags(session):
    return {c["id"]: [(f["phrase"].lower(), f["kind"]) for f in c["vague"]] for c in session.cases()["cases"]}


# --- intake sees the vague wording ---------------------------------------------------------
def test_intake_flags_a_result_judged_against_a_document_nobody_named(vague_session):
    found = flags(vague_session)
    assert ("đúng design", "reference") in found["V-1"] and not any(kind == "unmeasurable" and "hiển thị đúng" in p for p, kind in found["V-1"])
    assert found["V-2"] == [("correctly", "unmeasurable")]
    assert found["V-3"] == []
    assert found["V-4"] == [("(no expected result)", "unmeasurable")]
    assert found["V-5"] == [("attachment", "reference")]


@pytest.mark.parametrize("text,kind", [("Matches the design", "reference"), ("hiển thị theo thiết kế", "reference"), ("画面は仕様通りであること", "reference"), ("Page loads normally", "unmeasurable"),
                                       ("入力が正しく保存される", "unmeasurable")])
def test_the_phrases_in_english_vietnamese_and_japanese(text, kind):
    assert [f["kind"] for f in intake.find_vague({"expected": text})] == [kind]


def test_a_measurable_expectation_is_not_flagged():
    assert intake.find_vague({"expected": "The list shows 5 rows and the badge reads Confirmed"}) == []


def test_the_workbooks_hidden_parts_are_reported_not_ignored(tmp_path):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Cases"
    sheet.append(["No.", "Steps", "Expected Result"])
    sheet.append(["A-1", "1. Open", "Shows 3 rows"])
    sheet["C2"].comment = Comment("note for the tester", "x")
    book.create_sheet("Design notes").append(["layout", "two columns"])
    book.save(tmp_path / "w.xlsx")
    warnings = "\n".join(intake.build(tmp_path / "w.xlsx")["warnings"])
    assert "cell comment" in warnings and "Design notes" in warnings


# --- a ready case is made concrete, from a source that is not a guess --------------------------
def ready(session, tmp_path, **per_case):
    config = workspace.load("demo")
    data = filled(session.cases(), config)
    for case in data["cases"]:
        for step in case["steps"]:
            step["expected"] = "The form shows the customer field"
    for case in data["cases"]:
        case.update(per_case.get(case["id"], {}))
    return data, config


def test_a_ready_case_with_vague_wording_is_refused_until_each_phrase_is_made_concrete(vague_session, tmp_path):
    data, config = ready(vague_session, tmp_path)
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, vague_session.cases(), config)
    text = "\n".join(refused.value.problems)
    for phrase in ("đúng design", "correctly", "(no expected result)", "attachment"):
        assert f'"{phrase}"' in text
    assert "V-3" not in text


def test_a_guess_is_not_a_source_and_a_general_phrase_is_not_concrete(vague_session, tmp_path):
    data, config = ready(vague_session, tmp_path)
    for case in data["cases"]:
        for item in case["vague_resolved"]:
            item.update(concrete="the screen is displayed correctly", source="my own idea")
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, vague_session.cases(), config)
    text = "\n".join(refused.value.problems)
    assert "not another general phrase" in text and "needs its source" in text


def test_a_design_file_the_user_gave_makes_the_case_ready(vague_session, tmp_path, design):
    data, config = ready(vague_session, tmp_path)
    resolved = lambda phrase, concrete, source: {"phrase": phrase, "concrete": concrete, "source": source}
    by = {c["id"]: c for c in data["cases"]}
    by["V-1"].update(refs=[{"kind": "file", "text": "BD man hinh.xlsx", "status": "ok", "path": str(design), "note": ""}],
                     vague_resolved=[resolved("đúng design", "fields partner_id (Khach hang) then date_order (Ngay dat), in that order", "ref:BD man hinh.xlsx")])
    by["V-2"]["vague_resolved"] = [resolved("correctly", "the order saves and its status badge reads Draft", "domain")]
    by["V-4"]["vague_resolved"] = [resolved("(no expected result)", "the order is saved: the form leaves edit mode and shows the new name", "workbook")]
    by["V-5"].update(refs=[{"kind": "file", "text": "BD man hinh.xlsx", "status": "ok", "path": str(design), "note": ""}],
                     vague_resolved=[resolved("attachment", "the file named in the chatter appears in the attachment list", "ref:BD man hinh.xlsx")])
    analysis.validate(data, vague_session.cases(), config)                    # no SchemaError


def test_a_document_source_must_be_a_reference_that_is_really_there(vague_session, tmp_path):
    data, config = ready(vague_session, tmp_path)
    by = {c["id"]: c for c in data["cases"]}
    by["V-1"]["refs"] = [{"kind": "file", "text": "BD.xlsx", "status": "ok", "path": str(tmp_path / "nowhere.xlsx"), "note": ""}]
    by["V-1"]["vague_resolved"] = [{"phrase": "đúng design", "concrete": "fields partner_id then date_order in that order", "source": "ref:BD.xlsx"}]
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, vague_session.cases(), config)
    text = "\n".join(p for p in refused.value.problems if "V-1" in p)
    assert "there is no file at" in text


def test_a_step_may_not_state_its_expectation_in_the_same_vague_words(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config)
    data["cases"][0]["steps"][0]["expected"] = "The page works correctly"
    with pytest.raises(SchemaError, match="still vague"):
        analysis.validate(data, session.cases(), config)


def test_waiving_a_reference_and_giving_up_both_need_a_reason(session):
    config = workspace.load("demo")
    data = filled(session.cases(), config)
    by = {c["id"]: c for c in data["cases"]}
    by["ORD-003"]["refs"] = [{"kind": "file", "text": "import.xlsx", "status": "waived", "path": "", "note": ""}]
    by["ORD-002"].update(verdict="needs_user", questions=["How many?"], data_considered="")
    by["ORD-001"].update(verdict="blocked", reason="no data", data_considered="")
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, session.cases(), config)
    text = "\n".join(refused.value.problems)
    assert "waived: say in its note why" in text and "full sentence" in text and "needs data_considered" in text and "blocked needs data_considered" in text


# --- the user shares a file by its path: nothing is uploaded ------------------------------------
def test_a_path_in_free_text_is_found_with_spaces_quotes_and_trailing_punctuation(tmp_path, design):
    assert refs.extract_paths(f"Design nằm ở {design}, anh xem nhé.") == [str(design)]
    assert refs.extract_paths(f'file "{design}" và thư mục {design.parent}') == [str(design), str(design.parent)]
    assert refs.extract_paths(r"không có gì ở C:\khong\ton\tai\x.xlsx và /web") == []


def test_the_path_in_an_answer_becomes_a_reference_the_assistant_is_told_about(session, design):
    session.ask("Which design?", "ORD-001")
    session.answer(f"Design là {design}")
    assert [i["name"] for i in refs.items(session)] == [design.name]
    answered = [e for e in events.read(session.dir) if e["kind"] == "answered"][0]
    assert answered["files"] == [str(design)]
    assert "oqa ref show" in cli._describe(answered)


def test_a_path_that_does_not_exist_or_holds_a_secret_is_refused(session, tmp_path):
    with pytest.raises(refs.RefError, match="Not found"):
        refs.add(session, str(tmp_path / "missing.xlsx"))
    secret = session.dir / "auth" / "Admin.json"
    secret.parent.mkdir(parents=True, exist_ok=True)
    secret.write_text("{}")
    with pytest.raises(refs.RefError, match="never read"):
        refs.add(session, str(secret))


def test_ref_show_reads_a_workbook_as_text_and_ref_add_remembers_it(session, design, capsys):
    assert cli.main(["ref", "add", str(design), "--case", "ORD-001", "--ws", "demo", "--session", session.id]) == 0
    assert cli.main(["ref", "show", "BD man hinh.xlsx", "--ws", "demo", "--session", session.id]) == 0
    out = capsys.readouterr().out
    assert "partner_id | Khach hang | 1" in out and "hien thi truoc ngay dat" in out
    assert cli.main(["ref", "show", "unknown.xlsx", "--ws", "demo", "--session", session.id]) == 2


def test_the_template_after_an_answer_carries_it_and_finds_the_file_again(session, tmp_path_factory):
    other = tmp_path_factory.mktemp("elsewhere")
    (other / "import.xlsx").write_text("x")
    config = workspace.load("demo")
    assert [r["status"] for r in analysis.template(session.cases(), config, session)["cases"][2]["refs"]] == ["missing"]
    session.ask("Where is import.xlsx?")
    session.answer(f"It is in {other}")
    data = analysis.template(session.cases(), config, session)
    ref = data["cases"][2]["refs"][0]
    assert ref["status"] == "ok" and ref["path"] == str(other / "import.xlsx")
    assert data["_user_answers"][0]["answer"].startswith("It is in") and data["_shared_files"][0]["kind"] == "folder"


def test_the_dashboard_takes_pasted_paths_one_per_line_and_says_which_one_is_wrong(session, design, tmp_path):
    app = dashboard_server.Dashboard()
    where = {"ws": "demo", "id": session.id}
    status, body = app.handle("POST", "/api/refs", {}, {**where, "path": f"{design}\n{tmp_path / 'nope.pdf'}"})
    assert body["ok"] and body["added"] == [design.name] and "nope.pdf" in body["error"]
    status, body = app.handle("POST", "/api/refs", {}, {**where, "path": str(tmp_path / "nope.pdf")})
    assert "error" in body and "ok" not in body
    news = [e["kind"] for e in session.news()]
    assert "ref_added" in news
    assert app.handle("GET", "/api/session", {"ws": "demo", "id": session.id}, {})[1]["refs"][0]["name"] == design.name


# --- a case that needs the user holds the session, and the questions are asked at once ---------------
def test_needs_user_keeps_the_session_in_analyze_and_asks_everything_in_one_go(session, tmp_path):
    config = workspace.load("demo")
    data = filled(session.cases(), config, **{"ORD-002": {"verdict": "needs_user", "questions": ["How many orders should the list hold: more than 80?"],
                                                         "data_considered": "tried to create 80 orders, but the profile cannot create sales orders"}})
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", "demo", "--session", session.id]) == 0
    state = session.state()
    assert state["phase"] == "ANALYZE" and "ORD-002" in state["waiting"]["question"]
    assert cli.main(["strategy", "--ws", "demo", "--session", session.id]) == 2           # nothing moves on while the question is open
    session.answer("More than 80, the demo data has 120.")
    fixed = filled(session.cases(), config)
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "b.json", fixed)), "--ws", "demo", "--session", session.id]) == 0
    assert session.state()["phase"] == "CONNECT"


def test_all_ready_goes_straight_on(session, tmp_path):
    config = workspace.load("demo")
    assert cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", filled(session.cases(), config))), "--ws", "demo", "--session", session.id]) == 0
    assert session.state()["phase"] == "CONNECT" and session.state()["waiting"] is None


# --- helpers a script can use ---------------------------------------------------------------------------
def test_o_ref_names_the_missing_file_instead_of_inventing_one():
    from unittest.mock import MagicMock
    from oqa.odoo_ui import Odoo
    o = Odoo(MagicMock(), "http://x")
    with pytest.raises(AssertionError, match="ask the user for its path"):
        o.ref("BD.xlsx")
    o.resolve_ref = lambda name: "C:/files/BD.xlsx" if name == "BD.xlsx" else ""
    assert o.ref("BD.xlsx") == "C:/files/BD.xlsx"


def test_the_list_of_references_is_kept_out_of_the_assistants_hands(session, home):
    from test_hooks import guard
    assert guard("Write", home, file_path=str(session.dir / "refs.json")).returncode == 2
    assert guard("Read", home, file_path=str(session.dir / "strategy.md")).returncode == 0
