import json
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from oqa import analysis, intake, workspace  # noqa: E402
from oqa.secrets import MemoryStore  # noqa: E402
from oqa.state import Session  # noqa: E402


@pytest.fixture(autouse=True)
def home(tmp_path, monkeypatch):
    """Every test gets its own Home and never meets the real one, the real credential store or a pointer file."""
    monkeypatch.setenv("ODOO_TEST_HOME", str(tmp_path / "home"))
    monkeypatch.setenv("OQA_NO_EXPLORE_GATE", "1")             # sessions of the older tests may run a hand-written script; the gate has its own tests
    monkeypatch.setenv("OQA_NO_SERVICE", "1")                 # `oqa run` in a test runs in the test's own process, with the test's credential store
    monkeypatch.setenv("OQA_NO_OPEN", "1")                    # a test never opens a spreadsheet on the developer's screen
    monkeypatch.delenv("OQA_WS", raising=False)
    monkeypatch.delenv("OQA_SESSION", raising=False)
    monkeypatch.setattr("oqa.paths.pointer_file", lambda: tmp_path / "pointer.json")
    return tmp_path / "home"


@pytest.fixture
def store():
    return MemoryStore()


@pytest.fixture
def ws(store, tmp_path):
    (tmp_path / "refs").mkdir()
    return workspace.save({"name": "demo", "url": "http://localhost:8069", "profiles": [{"name": "Admin", "login": "admin"}, {"name": "Lead", "login": "lead"}],
                           "reference_dirs": [str(tmp_path / "refs")]}, store, {"Admin": "pw-admin", "Lead": "pw-lead"})


def make_workbook(path, rows, header=("No.", "Test Case Name", "Precondition", "Steps", "Expected Result")):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Cases"
    sheet.append(["Cover"])
    sheet.append(list(header))
    for row in rows:
        sheet.append(list(row))
    book.save(path)
    return path


@pytest.fixture
def workbook(tmp_path):
    return make_workbook(tmp_path / "cases.xlsx", [
        ("ORD-001", "Create order", "", "1. Open Sales\n2. Click New\n3. Save", "Order is saved"),
        ("ORD-002", "Pagination", "Needs 80 orders", "1. Open the list\n2. Go to page 2", "Page 2 shows rows"),
        ("ORD-003", "Import lines", "File import.xlsx exists", "1. Import import.xlsx\n2. Check lines", "Lines imported; continue from ORD-001"),
    ])


def filled(cases_doc, config, **per_case):
    """A complete, valid analysis built from the template."""
    data = analysis.template(cases_doc, config)
    for case in data["cases"]:
        case.update(profile=config["profiles"][0]["name"], verdict="ready", clarity={"score": 4, "notes": "clear"}, sufficiency={"score": 4, "missing": []},
                    edge_cases=["empty list"], reuse="",
                    refs=[dict(r, status="ok") if r["path"] else dict(r, status="waived", note="the case runs without this file in this test") for r in case["refs"]])
        for step in case["steps"]:
            step["expected"] = "ok"
        case.update(per_case.get(case["id"], {}))
    return data


@pytest.fixture
def session(ws, workbook):
    cases = intake.build(workbook, ws["reference_dirs"])
    return Session.create("demo", workbook, cases)


def write(path, data):
    Path(path).write_text(json.dumps(data), encoding="utf-8")
    return path
