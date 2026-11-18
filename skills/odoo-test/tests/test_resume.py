"""A failed step is repaired and run again ALONE, from the page where it stopped: the runner keeps the browser, the steps that passed and what they made. Against the fake Odoo."""
import threading
import time

import pytest

pytest.importorskip("playwright")

from test_run import approve, explain, fake, live, write_script  # noqa: E402,F401  (fixtures and helpers)
from oqa import events, odoo_ui, probe, runtime, scaffold, service, view  # noqa: E402
from oqa.state import GateError  # noqa: E402

@pytest.fixture(autouse=True)
def patient(monkeypatch):
    """These tests run a browser, a service thread and an HTTP server at once: a wait of 1.5 s (enough for test_run) is not, on a busy computer."""
    monkeypatch.setattr(odoo_ui, "TIMEOUT", 5000)


STEPS = '''
def step_1(t, s):
    o = t.o
    o.menu("Sales", "Orders")
    s.hl(".o_list_view")
    o.expect_visible(".o_list_view")


def step_2(t, s):
    o = t.o
    o.new()
    s.hl(".o_form_view")
    o.expect_visible(".o_form_view")


def step_3(t, s):
    o = t.o
    t.data["name"] = o.name("order1")
    o.fill("name", t.data["name"])
    o.m2o("partner_id", "Azure")
    s.hl("[name=name]", "[name=partner_id]")
    o.expect_value("name", t.data["name"])
    o.expect_value("partner_id", "Azure Interior")


def step_4(t, s):
    o = t.o
    o.save()
    o.expect_text("Saved")
    s.hl(".o_notification")
    s.equals(t.data["name"].startswith("ZZ_AUTOTEST_"), True, "the name step 3 made reached step 4")
'''
BROKEN_3 = STEPS.replace('o.m2o("partner_id", "Azure")', 'o.m2o("partner_id", "Nobody")')
BROKEN_2 = STEPS.replace("o.new()", 'o.page.click(".o_no_such_button")')


def attempts(session):
    return session.state()["attempts"].get("T-1", 0)


def test_a_repaired_step_runs_alone_from_the_page_where_it_stopped(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", BROKEN_3, data)
    runner = runtime.Runner(session, ws, store)
    with runtime_browser() as new_page:
        first = runner.run(["T-1"], new_page=new_page)[0]
        assert first["status"] == "FAIL" and first["failed_step"] == 3 and first["from_step"] == 1 and first["steps_passed"] == 2
        write_script(session, "T-1", STEPS, data)                             # the repair: only step 3 changed
        explain(session, "T-1")
        second = runner.run(["T-1"], new_page=new_page)[0]
    assert second["status"] == "PASS" and second["from_step"] == 3 and "picking up at step 3" in second["note"] and second["steps_passed"] == 4
    ran = [(e["attempt"], e["n"]) for e in events.read(session.dir) if e["kind"] == "step"]
    assert ran == [(1, 1), (1, 2), (1, 3), (2, 3), (2, 4)]                  # steps 1 and 2 were not run again
    started = [e.get("from_step") for e in events.read(session.dir) if e["kind"] == "case_started"]
    assert started == [None, 3]
    record = view.case_records(session)[0]
    assert [(s["n"], s["status"], s["carried"]) for s in record["attempts"][1]["steps"]] == [(1, "PASS", True), (2, "PASS", True), (3, "PASS", False), (4, "PASS", False)]
    assert record["attempts"][1]["from_step"] == 3 and session.state()["phase"] == "EXECUTE"


def runtime_browser():
    from oqa.browser import launch
    return launch(True)


def test_a_run_that_gets_further_does_not_use_up_a_repair(live):
    session, ws, store = live
    data = approve(session, ws, store)
    runner = runtime.Runner(session, ws, store)
    with runtime_browser() as new_page:
        write_script(session, "T-1", BROKEN_2, data)
        assert runner.run(["T-1"], new_page=new_page)[0]["failed_step"] == 2 and attempts(session) == 1          # the first run is always counted
        write_script(session, "T-1", BROKEN_3, data)
        explain(session, "T-1", 1)
        assert runner.run(["T-1"], new_page=new_page)[0]["failed_step"] == 3 and attempts(session) == 1          # step 3 is further than step 2: free
        explain(session, "T-1", 2)
        same = runner.run(["T-1"], new_page=new_page)[0]
        assert same["failed_step"] == 3 and attempts(session) == 2 and same["runs_left"] == 2                    # the same step again is a failed repair


def test_an_edited_earlier_step_sends_the_run_back_to_step_one(live):
    session, ws, store = live
    data = approve(session, ws, store)
    runner = runtime.Runner(session, ws, store)
    with runtime_browser() as new_page:
        write_script(session, "T-1", BROKEN_3, data)
        runner.run(["T-1"], new_page=new_page)
        write_script(session, "T-1", STEPS.replace('o.menu("Sales", "Orders")', 'o.menu("Sales", "Orders")\n    s.note = 1'), data)         # step 1 was edited as well
        explain(session, "T-1")
        again = runner.run(["T-1"], new_page=new_page)[0]
    assert again["from_step"] == 1 and "edited after they passed" in again["note"] and again["status"] == "PASS"


def test_start_at_a_step_needs_a_kept_page_and_costs_nothing_when_refused(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", STEPS, data)
    runner = runtime.Runner(session, ws, store)
    with runtime_browser() as new_page:
        with pytest.raises(GateError, match="cannot start at step 3"):
            runner.run(["T-1"], start=3, new_page=new_page)
    assert session.state()["attempts"] == {}


def test_fresh_forgets_what_was_kept(live):
    session, ws, store = live
    data = approve(session, ws, store)
    runner = runtime.Runner(session, ws, store)
    with runtime_browser() as new_page:
        write_script(session, "T-1", BROKEN_3, data)
        runner.run(["T-1"], new_page=new_page)
        write_script(session, "T-1", STEPS, data)
        explain(session, "T-1")
        assert runner.run(["T-1"], fresh=True, new_page=new_page)[0]["from_step"] == 1


def test_the_page_where_a_step_stopped_is_saved_and_can_be_searched_without_a_browser(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", BROKEN_3, data)
    result = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert result["failed_page"].endswith("failed_page.txt")
    found = probe.failed(session, "T-1", "partner_id")
    assert found["url"].startswith("http") and any("[name=partner_id]" in line for line in found["elements"]) and found["total"] > 3
    assert probe.failed(session, None, "Customer")["file"] == found["file"]                         # no case named: the latest failure of any case
    with pytest.raises(FileNotFoundError):
        probe.failed(session, "NOPE-1")


def test_the_controls_the_failed_step_was_after_are_named_in_the_run_output():
    lines = ['button.btn.o_form_button_save[name=save] "Save"', 'button.o_cp_action_menus.dropdown-toggle "Actions"', 'input[name=partner_id] "Customer"', 'a.o_menu_toggle "Sales"']
    reason = "TimeoutError: Locator.click: Timeout 8000ms exceeded. (- waiting for locator(\".o_cp_action_menus button\").filter(has_text=\"Print\").first)"
    assert probe.near(lines, reason) == ['button.o_cp_action_menus.dropdown-toggle "Actions"']
    assert probe.near(lines, "field name='partner_id' is empty") == ['input[name=partner_id] "Customer"']
    assert probe.near(lines, "no such thing") == []


def test_the_scaffold_writes_one_function_per_step_and_the_checks_read_them(live, tmp_path):
    session, ws, store = live
    approve(session, ws, store)
    scaffold.scaffold(session)
    path = scaffold.script_path(session, "T-1")
    text = path.read_text(encoding="utf-8")
    assert [f"def step_{n}(t, s):" in text for n in (1, 2, 3, 4)] == [True] * 4 and "t.data" in text and "with t.step" not in text
    assert scaffold.unwritten(path) == 4 and set(scaffold.step_hashes(path)) == {1, 2, 3, 4}
    path.write_text(STEPS.replace('    s.hl(".o_list_view")\n', "").replace("    o.expect_visible(\".o_form_view\")\n", ""), encoding="utf-8")
    assert scaffold.unhighlighted_steps(path) == [1] and scaffold.unchecked_steps(path) == [2]
    assert scaffold.step_hashes(path)[3] == scaffold.step_hashes(path)[3] != scaffold.step_hashes(path)[4]


def test_a_script_missing_a_step_function_fails_with_a_sentence(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", STEPS.replace("def step_4(t, s):", "def helper(t, s):"), data)
    result = runtime.Runner(session, ws, store).run(["T-1"])[0]
    assert result["status"] == "FAIL" and "no function for step(s) [4]" in result["reason"]


def test_the_runner_service_keeps_the_browser_between_runs(live):
    session, ws, store = live
    data = approve(session, ws, store)
    write_script(session, "T-1", BROKEN_3, data)
    box = service.RunnerService(session, ws, store, headless=True, idle=60)
    thread = threading.Thread(target=lambda: box.serve(service.announce_to(session)), daemon=True)
    thread.start()
    deadline = time.monotonic() + 20
    while not service.alive(session) and time.monotonic() < deadline:
        time.sleep(0.2)
    try:
        first = service.run(session, "live", ["T-1"])[0]
        assert first["status"] == "FAIL" and first["failed_step"] == 3
        write_script(session, "T-1", STEPS, data)
        explain(session, "T-1")
        second = service.run(session, "live", ["T-1"])[0]                    # a second request to the SAME browser: it picks up at step 3
        assert second["status"] == "PASS" and second["from_step"] == 3
        with pytest.raises(GateError, match="cannot start at step 9|not in the analysis|already passed"):
            service.run(session, "live", ["T-1"], start=9)
        assert service.alive(session)["pid"] > 0
    finally:
        assert service.stop(session)
        thread.join(timeout=20)
    assert not service.alive(session) and not thread.is_alive()


def test_the_service_starts_itself_in_the_background_and_goes_away_when_told(live):
    session, ws, store = live
    approve(session, ws, store)
    info = service.ensure(session, "live", False)                          # a real child process: the address is in runner.json, the child answers a ping
    try:
        assert info["pid"] != __import__("os").getpid() and service.alive(session)["headless"] is True
        assert service.ensure(session, "live", False)["pid"] == info["pid"]  # asked again: the same one, not a second
    finally:
        assert service.stop(session)
    assert not service.alive(session) and not service.pointer(session).exists()


# --- explore first: a script is built from code that already ran on the real page ------------------
import contextlib  # noqa: E402

from oqa import cli, recipe  # noqa: E402

PIECES = {
    1: 'o.menu("Sales", "Orders")\ns.hl(".o_list_view")\no.expect_visible(".o_list_view")',
    2: 'o.new()\ns.hl(".o_form_view")\no.expect_visible(".o_form_view")',
    3: 't.data["name"] = o.name("order1")\no.fill("name", t.data["name"])\no.m2o("partner_id", "Azure")\ns.hl("[name=name]", "[name=partner_id]")\n'
       'o.expect_value("name", t.data["name"])\no.expect_value("partner_id", "Azure Interior")',
    4: 'o.save()\no.expect_text("Saved")\ns.hl(".o_notification")\ns.equals(t.data["name"].startswith("ZZ_AUTOTEST_"), True, "the name step 3 made")',
}


@contextlib.contextmanager
def runner_service(session, ws, store):
    box = service.RunnerService(session, ws, store, headless=True, idle=120)
    thread = threading.Thread(target=lambda: box.serve(service.announce_to(session)), daemon=True)
    thread.start()
    deadline = time.monotonic() + 20
    while not service.alive(session) and time.monotonic() < deadline:
        time.sleep(0.2)
    try:
        yield
    finally:
        service.stop(session)
        thread.join(timeout=20)


def gated(session):
    with session.update() as state:
        state["explore"] = "required"


def tried(session, case, step, code, **more):
    return service.ask(session, "live", "try", {"case": case, "step": step, "code": code, **more})


def test_code_is_tried_on_the_real_page_and_only_what_worked_is_kept(live):
    session, ws, store = live
    approve(session, ws, store)
    with runner_service(session, ws, store):
        opened = service.ask(session, "live", "explore", {"case": "T-1"})
        assert opened["ok"] and opened["controls"] and "Apps" in " ".join(opened["controls"])
        bad = tried(session, "T-1", 1, 'o.menu("Sales", "Orderz")\ns.hl(".o_list_view")')                    # a guess that does not fit this page
        assert bad["ok"] is False and "Timeout" in bad["error"] and bad["controls"] and recipe.steps_of(recipe.load(session, "T-1")) == {}
        assert session.state()["attempts"] == {}                                                          # a wrong selector found out here costs nothing
        good = tried(session, "T-1", 1, PIECES[1])
        assert good["ok"] and good["pieces"] == 1 and good["checks"] >= 1 and any("o_list_button_add" in c for c in good["controls"])
        with pytest.raises(GateError, match="try step 2 before step 3"):
            tried(session, "T-1", 3, PIECES[3])
        with pytest.raises(GateError, match="Not exploring T-2"):
            tried(session, "T-2", 1, PIECES[1])
        assert tried(session, "T-1", 2, "s.hl('.o_no_such_thing')")["ok"] is False                           # s.hl has to find its element now
        assert tried(session, "T-1", 2, PIECES[2])["ok"] and set(recipe.steps_of(recipe.load(session, "T-1"))) == {1, 2}
        undone = tried(session, "T-1", 2, "", undo=True)
        assert undone["undone"] and set(recipe.steps_of(recipe.load(session, "T-1"))) == {1}


def test_explore_replays_the_verified_steps_to_reach_a_later_one(live):
    session, ws, store = live
    approve(session, ws, store)
    with runner_service(session, ws, store):
        service.ask(session, "live", "explore", {"case": "T-1"})
        for n in (1, 2):
            assert tried(session, "T-1", n, PIECES[n])["ok"]
        again = service.ask(session, "live", "explore", {"case": "T-1", "to": 3})                         # a new page, steps 1 and 2 played back
        assert again["ok"] and any("[name=partner_id]" in c for c in again["controls"])
        assert tried(session, "T-1", 3, PIECES[3])["ok"]
        with pytest.raises(GateError, match="step 4 has nothing verified"):
            service.ask(session, "live", "explore", {"case": "T-1", "to": 5})


def test_the_script_is_built_from_tried_steps_and_a_hand_made_one_is_refused(live, monkeypatch, capsys):
    session, ws, store = live
    data = approve(session, ws, store)
    gated(session)
    monkeypatch.setattr(cli, "default_store", lambda: store)
    flags = ["--ws", "live", "--session", session.id]
    write_script(session, "T-1", STEPS, data)                                                             # written from memory of another Odoo: no
    with pytest.raises(GateError, match="was not built from steps tried on the real page"):
        runtime.Runner(session, ws, store).run(["T-1"])
    assert session.state()["attempts"] == {}
    with runner_service(session, ws, store):
        service.ask(session, "live", "explore", {"case": "T-1"})
        for n in (1, 2, 3):
            assert tried(session, "T-1", n, PIECES[n])["ok"]
        assert cli.main(["build", "T-1", *flags]) == 2 and "has code that was not built from tried steps" in capsys.readouterr().err   # it will not overwrite a hand-written script
        assert cli.main(["build", "T-1", "--force", *flags]) == 0
        assert "not tried yet: step(s) [4]" in capsys.readouterr().out
        assert tried(session, "T-1", 4, PIECES[4])["ok"] and cli.main(["build", "T-1", *flags]) == 0 and "Next: oqa run T-1" in capsys.readouterr().out
    path = scaffold.script_path(session, "T-1")
    assert recipe.sealed(path.read_text(encoding="utf-8")) and path.read_text(encoding="utf-8").count("def step_") == 4
    assert runtime.Runner(session, ws, store).run(["T-1"])[0]["status"] == "PASS"                         # what ran once on the page runs again, with evidence
    path.write_text(path.read_text(encoding="utf-8").replace('"Azure"', '"Deco"'), encoding="utf-8")        # edited by hand afterwards
    with pytest.raises(GateError, match="was edited by hand after it was built"):
        runtime.Runner(session, ws, store).run(["T-1"])


def test_exploring_counts_as_looking_at_the_page_for_a_repair_note(live):
    session, ws, store = live
    approve(session, ws, store)
    session.begin_attempt("T-1")
    events.append(session.dir, "attempt_finished", case="T-1", attempt=1, status="FAIL", reason="TimeoutError: x", failed_step=1, script_sha="a")
    events.append(session.dir, "explore", case="T-1", step=1, ok=True)
    assert explain(session, "T-1")["probed"] is True
