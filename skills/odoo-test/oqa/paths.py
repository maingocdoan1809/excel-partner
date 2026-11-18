"""Where things live. Everything is under one Home folder; nothing is ever copied into the user's project.

Home = $ODOO_TEST_HOME, else the folder named in ~/.odoo-test.json ({"home": "..."}), else ~/.odoo-test.

    <home>/workspaces/<ws>/workspace.json
    <home>/workspaces/<ws>/knowledge/index.md
    <home>/workspaces/<ws>/current                 id of the session the commands use when none is named
    <home>/workspaces/<ws>/sessions/<id>/{state,cases,analysis,strategy}.json, strategy.md, events.jsonl, evidence/, logs/
"""
import json
import os
import re
import secrets
import time
from pathlib import Path

HOME_ENV, WS_ENV, SESSION_ENV = "ODOO_TEST_HOME", "OQA_WS", "OQA_SESSION"
NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,59}$")


class PathError(Exception):
    """A name or a lookup that cannot be resolved; the message says what to do."""


def pointer_file() -> Path:
    return Path.home() / ".odoo-test.json"


def default_home() -> Path:
    return Path.home() / ".odoo-test"


def home() -> Path:
    if os.environ.get(HOME_ENV):
        return Path(os.environ[HOME_ENV]).expanduser()
    try:
        chosen = json.loads(pointer_file().read_text(encoding="utf-8")).get("home")
        if isinstance(chosen, str) and chosen.strip():
            return Path(chosen).expanduser()
    except (OSError, ValueError, AttributeError):
        pass
    return default_home()


def set_home(folder: str) -> Path:
    """Remember a Home other than the default. Existing data is not moved."""
    target = Path(folder).expanduser()
    if not target.is_absolute():
        raise PathError("The storage folder must be an absolute path.")
    if target == default_home():
        pointer_file().unlink(missing_ok=True)
    else:
        pointer_file().write_text(json.dumps({"home": str(target)}), encoding="utf-8")
    return target


def check_name(name: str, what: str = "name") -> str:
    if not isinstance(name, str) or not NAME_RE.match(name):
        raise PathError(f"The {what} must be 1-60 letters, digits, '.', '_' or '-' and start with a letter or digit.")
    return name


def workspaces_dir() -> Path:
    return home() / "workspaces"


def workspace_dir(name: str) -> Path:
    return workspaces_dir() / check_name(name, "workspace name")


def workspace_names() -> list:
    base = workspaces_dir()
    return sorted(p.name for p in base.iterdir() if (p / "workspace.json").exists()) if base.is_dir() else []


def session_dir(ws: str, session_id: str) -> Path:
    return workspace_dir(ws) / "sessions" / check_name(session_id, "session id")


def session_ids(ws: str) -> list:
    base = workspace_dir(ws) / "sessions"
    return sorted(p.name for p in base.iterdir() if (p / "state.json").exists()) if base.is_dir() else []


def new_session_id() -> str:
    return f"{time.strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(2)}"


def resolve_workspace(name=None) -> str:
    """Flag, else $OQA_WS, else the only workspace there is."""
    name = name or os.environ.get(WS_ENV)
    if name:
        check_name(name, "workspace name")
        if not (workspace_dir(name) / "workspace.json").exists():
            raise PathError(f"Workspace '{name}' does not exist. Run: oqa init --ws {name}")
        return name
    names = workspace_names()
    if len(names) == 1:
        return names[0]
    raise PathError("No workspace was named. Run: oqa init" if not names else f"Several workspaces exist ({', '.join(names)}). Say which with --ws.")


def resolve_session(ws=None, session=None):
    """(workspace, session id): flags, else $OQA_WS / $OQA_SESSION, else the workspace's `current` session."""
    ws = resolve_workspace(ws)
    session = session or os.environ.get(SESSION_ENV)
    if not session:
        pointer = workspace_dir(ws) / "current"
        session = pointer.read_text(encoding="utf-8").strip() if pointer.exists() else ""
    if not session:
        raise PathError(f"No session is open in workspace '{ws}'. Run: oqa session new --file <testcase file>")
    if not (session_dir(ws, session) / "state.json").exists():
        raise PathError(f"Session '{session}' does not exist in workspace '{ws}'.")
    return ws, session
