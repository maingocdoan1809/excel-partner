"""The dashboard's API: history, one session's view, pictures (and only pictures), and the two things only the user does."""
import json
import queue
import threading
import urllib.error
import urllib.request

import pytest

from conftest import filled, write
from oqa import dashboard_server, events, view, workspace
from oqa.state import Session
from oqa.store import read_json, write_json


class Up:
    def __init__(self, select=None):
        self.app = dashboard_server.Dashboard(select, timeout=30)
        self.urls = queue.Queue()
        threading.Thread(target=lambda: self.app.serve(self.urls.put, open_browser=False, forever=True), daemon=True).start()
        self.base = self.urls.get(timeout=5).split("/?")[0]

    def call(self, path, body=None, token=None):
        sep = "&" if "?" in path else "?"
        request = urllib.request.Request(f"{self.base}{path}{sep}t={token or self.app.token}", data=None if body is None else json.dumps(body).encode(),
                                         method="GET" if body is None else "POST", headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                raw = response.read()
                return response.status, (json.loads(raw) if response.headers["Content-Type"].startswith("application/json") else raw)
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read() or b"{}")


@pytest.fixture
def up():
    return Up()


def at_confirm(session, tmp_path):
    config = workspace.load(session.ws)
    data = filled(session.cases(), config)
    write_json(session.dir / "analysis.json", write(tmp_path / "a.json", data) and data)
    session.analysis_submitted()
    session.connected({"profiles": {"Admin": {"landing_url": "x", "version": "17.0"}}, "version": "17.0"})
    doc, text = __import__("oqa.analysis", fromlist=["x"]).strategy(data, session.cases(), config)
    write_json(session.dir / "strategy.json", doc)
    (session.dir / "strategy.md").write_text(text, encoding="utf-8")
    session.strategy_ready(doc["sha"])


def test_history_and_view_are_served_with_the_token_only(session, up):
    assert up.call("/api/sessions")[1]["sessions"][0]["session"] == session.id
    status, body = up.call(f"/api/session?ws=demo&id={session.id}")
    assert status == 200 and body["phase"] == "ANALYZE" and body["total"] == 3 and [c["id"] for c in body["cases"]] == ["ORD-001", "ORD-002", "ORD-003"]
    assert up.call("/api/sessions", token="wrong")[0] == 403
    page = urllib.request.urlopen(f"{up.base}/?t={up.app.token}", timeout=5).read().decode()
    assert "__TOKEN__" not in page and up.app.token in page


def test_phase_track_marks_every_earlier_phase_done(session, tmp_path):
    at_confirm(session, tmp_path)
    states = {p["phase"]: p["state"] for p in view.session_view(session)["phases"]}
    assert states == {"ANALYZE": "done", "CONNECT": "done", "CONFIRM": "current", "EXECUTE": "todo", "REPORT": "todo", "DONE": "todo"}


def test_a_picture_is_served_but_nothing_outside_evidence(session, up):
    (session.dir / "evidence" / "A").mkdir(parents=True)
    (session.dir / "evidence" / "A" / "s1.png").write_bytes(b"\x89PNG" + b"0" * 20)
    (session.dir / "secret.png").write_bytes(b"\x89PNG" + b"1" * 20)
    ok = up.call(f"/api/image?ws=demo&id={session.id}&path=evidence/A/s1.png")
    assert ok[0] == 200 and ok[1].startswith(b"\x89PNG")
    for bad in ("secret.png", "evidence/../secret.png", "../../workspace.json", "evidence/A/s1.txt", "evidence"):
        assert up.call(f"/api/image?ws=demo&id={session.id}&path={bad}")[0] == 404, bad
    assert up.call(f"/api/image?ws=demo&id=..%2F..&path=x")[1].get("error")


def test_the_user_confirms_from_the_dashboard_and_a_tampered_strategy_is_refused(session, up, tmp_path):
    at_confirm(session, tmp_path)
    assert up.call(f"/api/session?ws=demo&id={session.id}")[1]["can_confirm"] is True
    doc = read_json(session.dir / "strategy.json")
    doc["order"] = doc["order"][::-1]
    write_json(session.dir / "strategy.json", doc)
    assert "strategy" in up.call("/api/confirm", {"ws": "demo", "id": session.id})[1]["error"]
    assert session.state()["phase"] == "CONFIRM"
    doc["order"] = doc["order"][::-1]
    write_json(session.dir / "strategy.json", doc)
    assert up.call("/api/confirm", {"ws": "demo", "id": session.id})[1] == {"ok": True}
    assert session.state()["phase"] == "EXECUTE" and session.state()["confirmation"]["by"] == "dashboard"


def test_the_user_answers_the_assistant_from_the_dashboard(session, up):
    assert up.call("/api/answer", {"ws": "demo", "id": session.id, "text": "x"})[1]["error"]          # nobody asked anything
    session.ask("Which customer?", "ORD-001")
    view_ = up.call(f"/api/session?ws=demo&id={session.id}")[1]
    assert view_["waiting"]["question"] == "Which customer?"
    assert up.call("/api/answer", {"ws": "demo", "id": session.id, "text": "  "})[1]["error"]
    assert up.call("/api/answer", {"ws": "demo", "id": session.id, "text": "Azure Interior"})[1] == {"ok": True}
    assert any(e["kind"] == "answered" and e["answer"] == "Azure Interior" for e in events.read(session.dir))


def test_case_statuses_follow_the_events(session, tmp_path):
    at_confirm(session, tmp_path)
    session.confirm_current("t")
    session.begin_attempt("ORD-001")
    events.append(session.dir, "case_started", case="ORD-001", attempt=1)
    assert {c["id"]: c["status"] for c in view.case_records(session)}["ORD-001"] == "RUNNING"
    events.append(session.dir, "attempt_finished", case="ORD-001", attempt=1, status="FAIL", reason="x")
    assert {c["id"]: c["status"] for c in view.case_records(session)}["ORD-001"] == "REPAIRING"
    session.finish_case("ORD-001", "BLOCKED", "needs data")
    records = {c["id"]: c for c in view.case_records(session)}
    assert records["ORD-001"]["status"] == "BLOCKED" and records["ORD-001"]["reason"] == "needs data" and records["ORD-002"]["status"] == "PLANNED"
    assert view.list_sessions()[0]["counts"]["BLOCKED"] == 1


def test_asking_for_the_dashboard_again_reuses_the_running_one_on_the_right_session(session, up, monkeypatch, capsys):
    from oqa import build_id, cli, paths
    write_json(paths.home() / "dashboard.json", {"url": f"{up.base}/?t={up.app.token}", "pid": 0, "build": build_id()})
    opened = []
    monkeypatch.setattr(cli.webbrowser, "open", opened.append)
    assert cli.main(["dashboard", "--ws", "demo", "--session", session.id]) == 0
    assert "already running" in capsys.readouterr().out
    assert opened == [f"{up.base}/?t={up.app.token}#demo/{session.id}"]          # a new tab on THIS session; no second server was started


def test_a_stale_dashboard_pointer_is_not_trusted(session, monkeypatch):
    from oqa import cli, paths
    write_json(paths.home() / "dashboard.json", {"url": "http://127.0.0.1:9/?t=x", "pid": 0})
    assert cli._alive("http://127.0.0.1:9/?t=x") is False


def test_the_user_stops_resumes_and_hints_from_the_dashboard(session, up, tmp_path):
    at_confirm(session, tmp_path)
    session.confirm_current("t")
    sid = {"ws": "demo", "id": session.id}
    view_ = up.call(f"/api/session?ws=demo&id={session.id}")[1]
    assert view_["can_stop"] is True and view_["stopped"] is None
    assert up.call("/api/stop", sid)[1] == {"ok": True}
    assert up.call("/api/stop", sid)[1]["error"]                                                  # already stopped
    assert up.call(f"/api/session?ws=demo&id={session.id}")[1]["stopped"]["by"] == "dashboard"
    assert up.call("/api/resume", sid)[1] == {"ok": True}
    session.begin_attempt("ORD-001")
    case = next(c for c in up.call(f"/api/session?ws=demo&id={session.id}")[1]["cases"] if c["id"] == "ORD-001")
    assert case["can_hint"] is True and case["exhausted"] is False
    assert "Write the hint" in up.call("/api/hint", {**sid, "case": "ORD-001", "text": " "})[1]["error"]
    assert up.call("/api/hint", {**sid, "case": "ORD-001", "text": "use the Actions menu"})[1] == {"ok": True}
    case = next(c for c in up.call(f"/api/session?ws=demo&id={session.id}")[1]["cases"] if c["id"] == "ORD-001")
    assert case["hints"][0]["text"] == "use the Actions menu" and case["used"] == 0


def test_a_case_that_used_its_budget_is_flagged_as_needing_the_user(session, tmp_path):
    at_confirm(session, tmp_path)
    session.confirm_current("t")
    for _ in range(4):
        session.begin_attempt("ORD-001")
    assert next(c for c in view.case_records(session) if c["id"] == "ORD-001")["exhausted"] is True


def test_the_user_can_build_and_download_the_excel_evidence_from_the_dashboard(session, up, tmp_path):
    at_confirm(session, tmp_path)
    sid = {"ws": "demo", "id": session.id}
    assert up.call(f"/api/session?ws=demo&id={session.id}")[1]["reports"] == []
    status, built = up.call("/api/report", sid)
    assert status == 200 and built["ok"] and built["main"].startswith("Evidence_") and built["main"].endswith(f"{session.id}.xlsx")
    shown = up.call(f"/api/session?ws=demo&id={session.id}")[1]["reports"]
    assert shown[0]["main"] is True and shown[0]["name"] == built["main"]
    got = urllib.request.urlopen(f"{up.base}/api/report?ws=demo&id={session.id}&name={built['main']}&t={up.app.token}", timeout=5)
    body = got.read()
    assert body[:2] == b"PK" and "spreadsheetml" in got.headers["Content-Type"] and built["main"] in got.headers["Content-Disposition"]
    for bad in ("../state.json", "SUMMARY.md/../../state.json", "nothing.xlsx"):
        assert up.call(f"/api/report?ws=demo&id={session.id}&name={bad}")[0] == 404, bad


def test_fonts_dropped_into_the_fonts_folder_are_served_to_the_page_without_a_token_but_nothing_else(up, tmp_path, monkeypatch):
    from oqa import localserver
    folder = tmp_path / "fonts"
    folder.mkdir()
    (folder / "BeVietnamPro-Regular.woff2").write_bytes(b"wOF2" + b"0" * 20)
    (folder / "notes.txt").write_text("not a font")
    (tmp_path / "secret.ttf").write_bytes(b"secret")
    monkeypatch.setattr(localserver, "FONTS", folder)

    def fetch(name, host=None):
        request = urllib.request.Request(f"{up.base}/font/{name}", headers={"Host": host} if host else {})
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, response.headers["Content-Type"], response.read()
        except urllib.error.HTTPError as error:
            return error.code, "", b""
    assert fetch("BeVietnamPro-Regular.woff2")[:2] == (200, "font/woff2")
    for bad in ("missing.woff2", "notes.txt", "..%2Fsecret.ttf", "%2e%2e/secret.ttf"):
        assert fetch(bad)[0] in (403, 404), bad
    assert fetch("BeVietnamPro-Regular.woff2", host="evil.example")[0] == 403           # only the page of this very server may use it


def test_a_dashboard_started_from_older_files_is_replaced_not_reused(up, capsys):
    from oqa import build_id, cli, paths
    assert len(build_id()) == 12 and build_id() == build_id()
    write_json(paths.home() / "dashboard.json", {"url": f"{up.base}/?t={up.app.token}", "pid": 0})            # a pointer from before builds were recorded
    assert cli.main(["dashboard", "--no-open", "--timeout", "2"]) == 0
    assert up.app.finished() is True                                                                       # the old one was asked to leave
    assert "Dashboard: http" in capsys.readouterr().out and not (paths.home() / "dashboard.json").exists()  # the new one announced itself, then ended and removed its pointer


def test_the_dashboard_that_matches_the_files_is_kept(up, capsys, monkeypatch):
    from oqa import build_id, cli, paths
    monkeypatch.setattr("webbrowser.open", lambda url: True)
    write_json(paths.home() / "dashboard.json", {"url": f"{up.base}/?t={up.app.token}", "pid": 0, "build": build_id()})
    assert cli.main(["dashboard"]) == 0 and "already running" in capsys.readouterr().out and up.app.finished() is False
