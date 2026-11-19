"""PreToolUse hook: stops the shortcuts around the process. Reads the tool call as JSON on stdin; exit 2 blocks it and the message goes back to the
assistant. The process itself is enforced inside oqa (every command checks the session's phase); this only closes the doors oqa cannot see:
editing its files by hand, reading signed-in cookies, confirming for the user, and running a case script without the runner's evidence checks."""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from oqa import paths  # noqa: E402

OWNED = re.compile(r"/sessions/[^/]+/(state\.json|events\.jsonl|cases\.json|refs\.json|analysis\.json|strategy\.json|strategy\.md|evidence/|report/|logs/)|/workspace\.json$|/knowledge/index\.md$|\.odoo-test\.json$")
SECRET = re.compile(r"/sessions/[^/]+/auth/|storageState|/credentials", re.I)
CONFIRM = re.compile(r"(^|[\s;&|\"'])(?:\S*oqa(?:\.py)?|-m\s+oqa)\s+(?:confirm|resume|stop)\b")
SCRIPT = re.compile(r"\bpython[\w.]*\s+(?:-\S+\s+)*\S*[/\\]workspaces[/\\][^\s]+[/\\]scripts[/\\]\S+\.py")
SECRET_SHELL = re.compile(r"(sessions[/\\][^\s\"']+[/\\]auth|storageState)", re.I)


def inside_home(path):
    try:
        return Path(path).resolve().is_relative_to(paths.home().resolve())
    except (OSError, ValueError):
        return False


def verdict(call):
    tool, args = call.get("tool_name", ""), call.get("tool_input") or {}
    target = str(args.get("file_path") or args.get("path") or "").replace("\\", "/")
    if tool in ("Write", "Edit", "MultiEdit", "NotebookEdit") and target and inside_home(target):
        if SECRET.search(target) or OWNED.search(target):
            return "This file belongs to oqa (session state, evidence, workspace or knowledge). Change it with an oqa command, not by hand: oqa knowledge learn, oqa analyze --submit, oqa ask / answer, oqa finish."
    if tool in ("Read", "Grep", "Glob") and target and SECRET.search(target):
        return "Signed-in cookies and passwords are never read: they are used by oqa only."
    if tool == "Bash":
        command = str(args.get("command", ""))
        if CONFIRM.search(command):
            return "Only the user confirms the strategy and stops or resumes the test: they do it on the dashboard or run the command in their own terminal. Tell them what is waiting, then `oqa wait` in the background."
        if SECRET_SHELL.search(command):
            return "Signed-in cookies and passwords are never read: they are used by oqa only."
        if SCRIPT.search(command):
            return "Run case scripts with `oqa run <case>`: running one directly skips the evidence checks and the attempt count."
    return None


if __name__ == "__main__":
    try:
        message = verdict(json.load(sys.stdin))
    except Exception:
        sys.exit(0)                 # a hook that cannot read its input must not stop the work
    if message:
        print(message, file=sys.stderr)
        sys.exit(2)
