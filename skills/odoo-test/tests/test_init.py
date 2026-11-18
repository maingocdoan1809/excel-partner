"""`oqa init`: the temporary form server."""
import http.server
import json
import queue
import threading
import urllib.error
import urllib.request

import pytest

from oqa import init_server, paths, workspace
from oqa.secrets import MemoryStore


class Running:
    def __init__(self, server):
        self.server, self.urls, self.result = server, queue.Queue(), []
        self.thread = threading.Thread(target=lambda: self.result.append(server.run(self.urls.put, open_browser=False, )), daemon=True)
        self.thread.start()
        self.url = self.urls.get(timeout=5)
        self.base = self.url.split("/?")[0]

    def call(self, path, body=None, token=None, headers=None):
        request = urllib.request.Request(self.base + path, data=None if body is None else json.dumps(body).encode(), method="GET" if body is None else "POST",
                                         headers={"X-OQA-Token": self.server.token if token is None else token, "Content-Type": "application/json", **(headers or {})})
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read() or b"{}")

    def finish(self):
        self.thread.join(5)
        return self.result[0] if self.result else "still running"


FORM = {"workspace": {"name": "erp", "url": "localhost:8069", "profiles": [{"name": "Admin", "login": "admin", "note": "all rights"}], "reference_dirs": ["C:/docs"]},
        "passwords": {"Admin": "s3cret"}}


def test_form_is_only_served_with_the_token_and_hides_nothing_secret(store):
    run = Running(init_server.InitServer(store, "erp", timeout=10))
    with pytest.raises(urllib.error.HTTPError) as refused:
        urllib.request.urlopen(run.base + "/?t=wrong", timeout=5)
    assert refused.value.code == 403
    page = urllib.request.urlopen(run.url, timeout=5).read().decode()
    assert run.server.token in page and "__TOKEN__" not in page and "__BOOT__" not in page
    assert run.call("/api/save", FORM)[1]["ok"]
    run.finish()


def test_api_refuses_a_wrong_token_a_foreign_origin_and_a_foreign_host(store):
    run = Running(init_server.InitServer(store, timeout=10))
    assert run.call("/api/check", {"url": "http://x"}, token="nope")[0] == 403
    assert run.call("/api/save", FORM, headers={"Origin": "http://evil.example"})[0] == 403
    assert run.call("/api/save", FORM, headers={"Host": "evil.example"})[0] == 403
    assert store.items == {}                                              # nothing was saved by any of them
    run.call("/api/save", FORM)
    run.finish()


def test_saving_writes_the_workspace_keeps_the_password_in_the_store_and_stops_the_server(store):
    run = Running(init_server.InitServer(store, timeout=10))
    status, body = run.call("/api/save", FORM)
    assert status == 200 and body == {"ok": True, "name": "erp"}
    summary = run.finish()
    assert summary["workspace"] == "erp" and summary["url"] == "http://localhost:8069" and "s3cret" not in json.dumps(summary)
    assert store.get("oqa:erp:Admin") == "s3cret"
    assert "s3cret" not in (paths.workspace_dir("erp") / "workspace.json").read_text(encoding="utf-8")
    assert workspace.load("erp")["reference_dirs"] == ["C:/docs"]


def test_a_bad_form_gets_every_problem_back_and_the_server_keeps_waiting(store):
    run = Running(init_server.InitServer(store, timeout=10))
    status, body = run.call("/api/save", {"workspace": {"name": "x y", "url": "", "profiles": []}})
    assert status == 200 and "URL" in body["error"] and "Profiles" in body["error"]
    assert run.thread.is_alive()
    run.call("/api/save", FORM)
    run.finish()


def test_reopening_shows_the_saved_values_and_never_the_password(store):
    workspace.save(FORM["workspace"], store, FORM["passwords"])
    run = Running(init_server.InitServer(store, "erp", timeout=10))
    page = urllib.request.urlopen(run.url, timeout=5).read().decode()
    assert '"has_secret": true' in page and "s3cret" not in page and "admin" in page
    status, body = run.call("/api/workspace?name=erp")
    assert body["workspace"]["profiles"][0]["has_secret"] is True and "s3cret" not in json.dumps(body)
    run.call("/api/save", {"workspace": dict(FORM["workspace"], url="http://localhost:9999")})        # blank password = keep the saved one
    run.finish()
    assert workspace.load("erp")["url"] == "http://localhost:9999" and store.get("oqa:erp:Admin") == "s3cret"


def test_nothing_is_written_when_the_time_runs_out(store):
    assert init_server.InitServer(store, timeout=0.5).run(lambda url: None, open_browser=False) is None
    assert paths.workspace_names() == []


def test_the_storage_folder_can_be_changed_unless_the_environment_fixes_it(store, tmp_path, monkeypatch):
    run = Running(init_server.InitServer(store, timeout=10))
    status, body = run.call("/api/save", dict(FORM, home=str(tmp_path / "elsewhere")))
    assert "ODOO_TEST_HOME" in body["error"]
    monkeypatch.delenv("ODOO_TEST_HOME")
    status, body = run.call("/api/save", dict(FORM, home=str(tmp_path / "elsewhere")))
    run.finish()
    assert body["ok"] and (tmp_path / "elsewhere" / "workspaces" / "erp" / "workspace.json").exists()


def test_status_without_a_server(store):
    assert init_server.status()["configured"] is False
    workspace.save(FORM["workspace"], store, FORM["passwords"])
    assert init_server.status()["configured"] is True


def test_check_url_recognises_odoo_and_reports_a_dead_address():
    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"<html class='o_login'>Odoo</html>")

        def log_message(self, *args):
            pass

    server = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        found = init_server.check_url(f"http://127.0.0.1:{server.server_address[1]}")
        assert found["ok"] and found["odoo"]
    finally:
        server.shutdown()
    assert init_server.check_url("http://127.0.0.1:9", timeout=2)["ok"] is False
    assert init_server.check_url("ftp://x")["ok"] is False
