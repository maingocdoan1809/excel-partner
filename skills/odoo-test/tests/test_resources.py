"""Resource files: the declarations are checked, the files are made before the user is asked, the user sees and may change them, and what runs is what they approved."""
import base64
import io
import json
from types import SimpleNamespace
from unittest.mock import MagicMock

import openpyxl
import pytest

from conftest import filled, write
from oqa import analysis, cli, events, odoo_ui, resources, runtime, scaffold, view, workspace
from oqa.state import GateError, Session
from oqa.store import SchemaError, read_json
from test_dashboard import Up
from hooks import guard

SPEC = {"columns": ["Product", "Qty"], "rows": [["ZZ_AUTOTEST_p-{i}", 2]], "count": 3}
GUESSED = {"name": "import.xlsx", "mode": "guessed", "purpose": "the lines the import case loads", "spec": SPEC}
PICTURE = {"name": "photo.png", "mode": "described", "purpose": "the product picture", "description": "a blue square that says Chair", "spec": {"width": 80, "height": 60, "text": "Chair"}}
OWN = {"name": "price list.csv", "mode": "provided", "purpose": "the customer's own price list"}


@pytest.fixture
def up():
    return Up()


def declared(session, *items, uses=("import.xlsx",)):
    """A valid analysis whose ORD-003 declares `items`; its first step uses `uses`."""
    data = filled(session.cases(), workspace.load(session.ws))
    data["cases"][2]["resources"] = [dict(i) for i in items]
    data["cases"][2]["steps"][0]["uses"] = list(uses)
    return data


def submit(session, tmp_path, data):
    return cli.main(["analyze", "--submit", str(write(tmp_path / "a.json", data)), "--ws", session.ws, "--session", session.id])


def at_confirm(session, tmp_path, *items, uses=("import.xlsx",)):
    data = declared(session, *items, uses=uses)
    assert submit(session, tmp_path, data) == 0
    session.connected({"profiles": ["Admin"]})
    assert cli.main(["strategy", "--ws", session.ws, "--session", session.id]) == 0
    return data


def ask(up, session, case="ORD-003", **fields):
    return up.call("/api/resource", {"ws": session.ws, "id": session.id, "case": case, **fields})[1]


def b64(data):
    return base64.b64encode(data).decode()


def sheet_rows(path):
    return [list(r) for r in openpyxl.load_workbook(path).active.iter_rows(values_only=True)]


# --- making the files ----------------------------------------------------------------------
def test_a_spreadsheet_is_made_from_its_spec_and_repeats_its_rows_to_the_count():
    book = openpyxl.load_workbook(io.BytesIO(resources.render("xlsx", SPEC)))
    assert [list(r) for r in book.active.iter_rows(values_only=True)] == [["Product", "Qty"], ["ZZ_AUTOTEST_p-1", 2], ["ZZ_AUTOTEST_p-2", 2], ["ZZ_AUTOTEST_p-3", 2]]
    assert resources.render("csv", {"columns": ["a", "b"], "rows": [["x", 1]], "delimiter": ";"}) == b"a;b\nx;1\n"


def test_a_picture_a_document_and_a_text_are_made_too():
    from PIL import Image
    assert Image.open(io.BytesIO(resources.render("png", {"width": 90, "height": 50}))).size == (90, 50)
    assert resources.render("jpg", {})[:2] == b"\xff\xd8"
    assert resources.render("pdf", {"text": "hello " * 40, "pages": 2})[:4] == b"%PDF"
    assert resources.render("txt", {"text": "line"}) == b"line"


@pytest.mark.parametrize("kind, spec, word", [
    ("xlsx", {"columns": ["a", "a"]}, "different"),
    ("xlsx", {"columns": ["a"], "rows": [["x", "too many"]]}, "values each"),
    ("xlsx", {"columns": ["a"], "count": 5}, "at least one row"),
    ("png", {"width": 1}, "from 16 to 4000"),
    ("png", {"color": "blue"}, "#4a90d9"),
    ("docx", {}, "from the user"),
    ("xlsx", "not an object", "object"),
])
def test_a_spec_that_cannot_be_made_says_why(kind, spec, word):
    with pytest.raises(SchemaError) as refused:
        resources.render(kind, spec)
    assert word in "\n".join(refused.value.problems)


# --- the analysis --------------------------------------------------------------------------
def problems_of(session, *items, uses=("import.xlsx",), steps=None):
    data = declared(session, *items, uses=uses)
    if steps:
        data["cases"][2]["steps"][1]["uses"] = steps
    with pytest.raises(SchemaError) as refused:
        analysis.validate(data, session.cases(), workspace.load(session.ws))
    return "\n".join(refused.value.problems)


def test_the_analysis_must_use_every_resource_and_only_declared_ones(session):
    assert "uses 'ghost.xlsx', which is not a resource" in problems_of(session, GUESSED, uses=("import.xlsx", "ghost.xlsx"))
    assert "photo.png is not used by any step" in problems_of(session, GUESSED, PICTURE)


def test_a_resource_needs_a_good_name_a_mode_a_purpose_and_a_spec_that_can_be_made(session):
    assert "needs a `name`" in problems_of(session, {**GUESSED, "name": "../x.xlsx"})
    assert "declared twice" in problems_of(session, GUESSED, GUESSED)
    assert "mode must be one of" in problems_of(session, {**GUESSED, "mode": "invented"})
    assert "what it is for" in problems_of(session, {**GUESSED, "purpose": " "})
    assert "different" in problems_of(session, {**GUESSED, "spec": {"columns": ["a", "a"]}})
    assert "the user's own words" in problems_of(session, {**PICTURE, "description": ""}, uses=("photo.png",))
    assert "source file not found" in problems_of(session, {**OWN, "source": "/no/such/file.csv"}, uses=("price list.csv",))
    assert "cannot be made from a description" in problems_of(session, {**GUESSED, "name": "contract.docx"}, uses=("contract.docx",))


def test_the_template_has_room_for_resources_and_a_case_without_them_is_still_valid(session):
    template = analysis.template(session.cases(), workspace.load(session.ws))
    assert all(c["resources"] == [] and all(s["uses"] == [] for s in c["steps"]) for c in template["cases"])
    analysis.validate(filled(session.cases(), workspace.load(session.ws)), session.cases(), workspace.load(session.ws))


# --- before the user approves --------------------------------------------------------------
def test_accepting_the_analysis_makes_the_files_and_the_strategy_lists_them(session, tmp_path, capsys):
    at_confirm(session, tmp_path, GUESSED)
    assert sheet_rows(resources.file_of(session.dir, "ORD-003", "import.xlsx"))[1] == ["ZZ_AUTOTEST_p-1", 2]
    entry = resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]
    assert entry["mode"] == "guessed" and entry["facts"]["rows"] == 3 and entry["facts"]["columns"] == ["Product", "Qty"] and not entry["edited"]
    assert "import.xlsx [AI đoán]" in (session.dir / "strategy.md").read_text(encoding="utf-8")
    assert cli.main(["resource", "list", "--ws", session.ws, "--session", session.id]) == 0
    assert "ORD-003: import.xlsx" in capsys.readouterr().out
    assert cli.main(["resource", "show", "ORD-003", "import.xlsx", "--ws", session.ws, "--session", session.id]) == 0
    assert "ZZ_AUTOTEST_p-2 | 2" in capsys.readouterr().out
    assert cli.main(["resource", "show", "ORD-003", "nope.xlsx", "--ws", session.ws, "--session", session.id]) == 2


def test_a_file_the_user_must_give_blocks_the_approval_until_it_is_uploaded(session, tmp_path, up):
    at_confirm(session, tmp_path, OWN, uses=("price list.csv",))
    state = up.call(f"/api/session?ws=demo&id={session.id}")[1]
    assert state["resources_blocking"] == ["ORD-003: price list.csv needs the user's file: upload it on the dashboard"]
    assert "price list.csv" in up.call("/api/confirm", {"ws": "demo", "id": session.id})[1]["error"]
    with pytest.raises(GateError, match="Not every resource"):
        session.confirm_current("terminal")
    assert ask(up, session, name="price list.csv", filename="mine.csv", data=b64(b"Name,Price\nZZ_AUTOTEST_a,5\n")) == {"ok": True}
    state = up.call(f"/api/session?ws=demo&id={session.id}")[1]
    assert state["resources_blocking"] == []
    shown = next(r for k in state["cases"] for r in k["resources"] if k["id"] == "ORD-003")
    assert shown["mode"] == "provided" and shown["edited"] and shown["source"] == "mine.csv" and shown["facts"]["rows"] == 1


def test_the_user_may_change_a_guess_and_the_assistant_cannot_undo_it(session, tmp_path, up):
    data = at_confirm(session, tmp_path, GUESSED)
    changed = {**SPEC, "count": 5}
    assert ask(up, session, name="import.xlsx", spec=changed) == {"ok": True}
    entry = resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]
    assert entry["mode"] == "described" and entry["edited"] and entry["facts"]["rows"] == 5
    assert submit(session, tmp_path, data) == 0                                    # the assistant hands the analysis in again, with its own guess
    again = resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]
    assert again["spec"] == changed and again["edited"] and again["facts"]["rows"] == 5


def test_a_resubmitted_analysis_keeps_an_unchanged_guess_and_drops_a_resource_that_is_gone(session, tmp_path):
    data = at_confirm(session, tmp_path, GUESSED, PICTURE, uses=("import.xlsx", "photo.png"))
    before = resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]["sha256"]
    assert submit(session, tmp_path, data) == 0
    assert resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]["sha256"] == before
    data["cases"][2]["resources"], data["cases"][2]["steps"][0]["uses"] = [GUESSED], ["import.xlsx"]
    assert submit(session, tmp_path, data) == 0
    assert not resources.file_of(session.dir, "ORD-003", "photo.png").exists() and "photo.png" not in resources.load(session.dir)["resources"]["ORD-003"]


def test_a_bad_edit_is_refused_and_changes_nothing(session, tmp_path, up):
    at_confirm(session, tmp_path, GUESSED)
    sha = resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]["sha256"]
    assert "different" in ask(up, session, name="import.xlsx", spec={"columns": ["a", "a"]})["error"]
    assert ".xlsx" in ask(up, session, name="import.xlsx", filename="mine.csv", data=b64(b"x"))["error"]
    assert "could not be read" in ask(up, session, name="import.xlsx", filename="a.xlsx", data="***")["error"]
    assert "empty" in ask(up, session, name="import.xlsx", filename="a.xlsx", data="")["error"]
    assert "no resource" in ask(up, session, name="other.xlsx", spec=SPEC)["error"]
    assert resources.load(session.dir)["resources"]["ORD-003"]["import.xlsx"]["sha256"] == sha


def test_a_file_without_the_test_prefix_is_flagged(session, tmp_path, up):
    at_confirm(session, tmp_path, GUESSED)
    plain = {"columns": ["Product"], "rows": [["Chair"]]}
    assert ask(up, session, name="import.xlsx", spec=plain) == {"ok": True}
    shown = next(r for k in up.call(f"/api/session?ws=demo&id={session.id}")[1]["cases"] for r in k["resources"])
    assert shown["warnings"] == ["prefix"]


def test_the_dashboard_serves_the_file_and_its_first_rows_but_only_declared_ones(session, tmp_path, up):
    at_confirm(session, tmp_path, GUESSED)
    status, body = up.call(f"/api/resource?ws=demo&id={session.id}&case=ORD-003&name=import.xlsx&what=preview")
    assert status == 200 and body["columns"] == ["Product", "Qty"] and body["total"] == 3 and body["rows"][0] == ["ZZ_AUTOTEST_p-1", "2"]
    status, raw = up.call(f"/api/resource?ws=demo&id={session.id}&case=ORD-003&name=import.xlsx&download=1")
    assert status == 200 and raw[:2] == b"PK"
    for name in ("../../workspace.json", "other.xlsx", ".."):
        assert up.call(f"/api/resource?ws=demo&id={session.id}&case=ORD-003&name={name}")[0] == 404, name
    assert up.call(f"/api/resource?ws=demo&id={session.id}&case=ORD-001&name=import.xlsx")[0] == 404


# --- the approval --------------------------------------------------------------------------
def test_the_approval_covers_the_files_the_user_looked_at(session, tmp_path, up):
    at_confirm(session, tmp_path, GUESSED)
    seen = up.call(f"/api/session?ws=demo&id={session.id}")[1]
    assert seen["resources_guessed"] == 1 and seen["resources_total"] == 1 and seen["resources_editable"] is True
    ask(up, session, name="import.xlsx", spec={**SPEC, "count": 4})                        # changed after they looked
    assert "changed since you last looked" in up.call("/api/confirm", {"ws": "demo", "id": session.id, "resources_sha": seen["resources_sha"]})[1]["error"]
    current = up.call(f"/api/session?ws=demo&id={session.id}")[1]["resources_sha"]
    assert up.call("/api/confirm", {"ws": "demo", "id": session.id, "resources_sha": current})[1] == {"ok": True}
    assert session.state()["confirmation"]["resources_sha"] == current and any(e.get("resources_sha") == current for e in events.read(session.dir) if e["kind"] == "confirmed")


def test_after_the_approval_the_files_can_no_longer_be_changed(session, tmp_path, up):
    at_confirm(session, tmp_path, GUESSED)
    session.confirm_current("terminal")
    assert "Not allowed in phase EXECUTE" in ask(up, session, name="import.xlsx", spec=SPEC)["error"]


# --- what a script gets --------------------------------------------------------------------
def confirmed(session, tmp_path, *items, uses=("import.xlsx",)):
    data = at_confirm(session, tmp_path, *items, uses=uses)
    session.confirm_current("terminal")
    return data["cases"][2]["steps"]


def step_of(session, spec, n=1):
    return runtime.Step(SimpleNamespace(spec=spec, session=session, case="ORD-003", checks=0), n)


def test_a_step_gets_the_file_it_named_with_what_is_inside_it(session, tmp_path):
    spec = confirmed(session, tmp_path, GUESSED)
    step = step_of(session, spec)
    f = step.fixture("import.xlsx")
    assert f.rows == 3 and f.columns == ["Product", "Qty"] and f.kind == "xlsx" and f.mode == "guessed" and str(f) == f.path
    assert sheet_rows(f.path)[0] == ["Product", "Qty"]
    assert step.used == [f.record()] and step.fixture("import.xlsx") and len(step.used) == 1


def test_a_step_cannot_use_a_file_it_did_not_name_or_one_that_changed_after_the_approval(session, tmp_path):
    spec = confirmed(session, tmp_path, GUESSED)
    with pytest.raises(runtime.StepFailed, match="does not name import.xlsx"):
        step_of(session, spec, 2).fixture("import.xlsx")
    with pytest.raises(runtime.StepFailed, match="has no resource"):
        step_of(session, [{**spec[0], "uses": ["ghost.xlsx"]}]).fixture("ghost.xlsx")
    resources.file_of(session.dir, "ORD-003", "import.xlsx").write_bytes(b"swapped")
    with pytest.raises(runtime.StepFailed, match="changed"):
        step_of(session, spec).fixture("import.xlsx")


def test_resources_edited_behind_the_users_back_stop_the_steps(session, tmp_path):
    spec = confirmed(session, tmp_path, GUESSED)
    path = session.dir / "resources.json"
    data = read_json(path)
    data["resources"]["ORD-003"]["import.xlsx"]["spec"] = {**SPEC, "count": 99}
    path.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(runtime.StepFailed, match="approved"):
        step_of(session, spec).fixture("import.xlsx")


def test_upload_attaches_the_file_to_a_file_input_or_through_the_chooser():
    page = MagicMock()
    o = odoo_ui.Odoo(page, "http://x")
    o.ready = lambda: o
    page.locator.return_value.first.evaluate.return_value = True
    o.upload("input[type=file]", SimpleNamespace(path="/tmp/a.xlsx"))
    page.locator.return_value.first.set_input_files.assert_called_once_with("/tmp/a.xlsx")
    page.locator.return_value.first.evaluate.return_value = False
    o.upload(".o_attach", "/tmp/b.png")
    page.expect_file_chooser.return_value.__enter__.return_value.value.set_files.assert_called_once_with("/tmp/b.png")


# --- what the dashboard and the report say -------------------------------------------------
def test_the_case_record_and_its_steps_say_which_files_they_use(session, tmp_path):
    spec = confirmed(session, tmp_path, GUESSED)
    record = next(k for k in view.case_records(session) if k["id"] == "ORD-003")
    assert [r["name"] for r in record["resources"]] == ["import.xlsx"] and record["resources"][0]["used_by"] == [1]
    assert record["planned"][0]["uses"] == ["import.xlsx"] and record["planned"][1]["uses"] == []
    f = step_of(session, spec).fixture("import.xlsx")
    events.append(session.dir, "case_started", case="ORD-003", attempt=1)
    events.append(session.dir, "step", case="ORD-003", attempt=1, n=1, status="PASS", caption="c", expected="e", actual="ok", images=[{"raw": "a", "annotated": "b"}], resources=[f.record()])
    first = next(k for k in view.case_records(session) if k["id"] == "ORD-003")["attempts"][0]["steps"][0]
    assert first["resources"][0]["name"] == "import.xlsx" and first["uses"] == ["import.xlsx"]


def test_the_scaffold_tells_the_assistant_where_each_file_goes(session, tmp_path):
    confirmed(session, tmp_path, GUESSED)
    script = scaffold.script_path(session, "ORD-003")
    scaffold.scaffold(session, ["ORD-003"])
    text = script.read_text(encoding="utf-8")
    assert 'in step 1: f = s.fixture("import.xlsx"); o.upload(<file input>, f)' in text and " upload " in text


def test_the_excel_report_names_the_files_and_who_decided_them(session, tmp_path):
    from oqa import report
    confirmed(session, tmp_path, GUESSED)
    built = report.build(session, workspace.load(session.ws))
    sheet = openpyxl.load_workbook(built["main"])["ORD-003"]
    text = "\n".join(str(c.value) for row in sheet.iter_rows() for c in row if c.value)
    assert "import.xlsx [AI tự đoán]" in text


def test_the_assistant_cannot_write_the_resources_by_hand(home):
    base = str(home / "workspaces" / "demo" / "sessions" / "s1")
    for name in ("resources.json", "fixtures/ORD-003/import.xlsx"):
        assert guard.verdict({"tool_name": "Write", "tool_input": {"file_path": f"{base}/{name}"}}), name
