"""The hooks, run the way Claude Code runs them: a subprocess fed the tool call as JSON."""
import json
import os
import subprocess
import sys
from pathlib import Path

from conftest import filled
from oqa import analysis, workspace
from oqa.store import write_json

ROOT = Path(__file__).resolve().parents[1]


def run(script, payload, home, **env):
    return subprocess.run([sys.executable, str(ROOT / "hooks" / script)], input=json.dumps(payload), capture_output=True, text=True,
                          env={**os.environ, "ODOO_TEST_HOME": str(home), "USERPROFILE": str(home), "HOME": str(home), **env})


def guard(tool, home, **args):
    return run("guard.py", {"tool_name": tool, "tool_input": args}, home)


def test_session_files_credentials_and_confirming_are_off_limits(session, home):
    state = str(session.dir / "state.json")
    for tool, args in (("Write", {"file_path": state}), ("Edit", {"file_path": str(session.dir / "events.jsonl")}), ("Write", {"file_path": str(session.dir / "evidence" / "A" / "s1.png")}),
                       ("Write", {"file_path": str(home / "workspaces" / "demo" / "workspace.json")}), ("Read", {"file_path": str(session.dir / "auth" / "Admin.json")}),
                       ("Bash", {"command": "python oqa.py confirm"}), ("Bash", {"command": "cd x && python -m oqa confirm --ws demo"}),
                       ("Bash", {"command": f"cat {session.dir}/auth/Admin.json"}),
                       ("Bash", {"command": f"python {home}/workspaces/demo/scripts/orders/T-1.py"})):
        result = guard(tool, home, **args)
        assert result.returncode == 2 and result.stderr.strip(), (tool, args)


def test_everything_else_is_allowed(session, home, tmp_path):
    for tool, args in (("Write", {"file_path": str(home / "workspaces" / "demo" / "scripts" / "orders" / "T-1.py")}), ("Write", {"file_path": str(tmp_path / "analysis.json")}),
                       ("Read", {"file_path": str(session.dir / "strategy.md")}), ("Bash", {"command": "python oqa.py run"}), ("Bash", {"command": "python oqa.py status --json"}),
                       ("Bash", {"command": "echo confirm the strategy with the user"}), ("Write", {"file_path": "C:/project/src/app.py"})):
        assert guard(tool, home, **args).returncode == 0, (tool, args)


def test_garbage_input_never_blocks(home):
    result = subprocess.run([sys.executable, str(ROOT / "hooks" / "guard.py")], input="not json", capture_output=True, text=True, env={**os.environ, "ODOO_TEST_HOME": str(home)})
    assert result.returncode == 0


def executing(session, tmp_path):
    config = workspace.load(session.ws)
    data = filled(session.cases(), config)
    write_json(session.dir / "analysis.json", data)
    session.analysis_submitted()
    session.connected({})
    doc, text = analysis.strategy(data, session.cases(), config)
    write_json(session.dir / "strategy.json", doc)
    session.strategy_ready(doc["sha"])
    session.confirm_current("t")


def stop(session, home):
    return run("stop.py", {"hook_event_name": "Stop"}, home, OQA_WS=session.ws, OQA_SESSION=session.id)


def test_stop_pushes_back_twice_then_lets_go(session, home, tmp_path):
    assert stop(session, home).returncode == 0                   # still analysing: stopping is fine
    executing(session, tmp_path)
    first, second, third = stop(session, home), stop(session, home), stop(session, home)
    assert (first.returncode, second.returncode, third.returncode) == (2, 2, 0)
    assert "ORD-001" in first.stderr and "1/2" in first.stderr and "2/2" in second.stderr


def test_stop_resets_when_there_is_progress_and_is_fine_when_the_user_is_asked(session, home, tmp_path):
    executing(session, tmp_path)
    assert stop(session, home).returncode == 2
    session.begin_attempt("ORD-001")
    session.finish_case("ORD-001", "BLOCKED", "x")
    after = stop(session, home)
    assert after.returncode == 2 and "1/2" in after.stderr                                      # progress made: the count starts again
    session.ask("Which customer?")
    assert stop(session, home).returncode == 0                                                 # waiting for the user: the assistant may stop


def test_the_assistant_cannot_stop_or_resume_the_test_for_the_user(session, home):
    for command in ("python oqa.py stop", "python oqa.py resume --ws demo", "cd x && python -m oqa resume"):
        assert guard("Bash", home, command=command).returncode == 2, command
    assert guard("Bash", home, command="python oqa.py wait").returncode == 0 and guard("Bash", home, command="python oqa.py status").returncode == 0


def test_a_stopped_session_may_be_left_alone(session, home, tmp_path):
    executing(session, tmp_path)
    session.stop("dashboard")
    assert stop(session, home).returncode == 0
