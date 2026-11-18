"""Stop hook: do not let the assistant stop while the session is executing cases and nobody is waiting for the user. It pushes back at most twice
for the same progress; the third time it lets go (a stuck assistant must end, not loop), and says so in the session log."""
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from oqa import paths  # noqa: E402
from oqa.state import Session  # noqa: E402
from oqa.store import read_json, write_json  # noqa: E402

MAX_PUSHBACKS = 2


def verdict():
    ws, session_id = paths.resolve_session(os.environ.get(paths.WS_ENV), os.environ.get(paths.SESSION_ENV))
    session = Session(ws, session_id)
    state = session.state()
    if state["phase"] != "EXECUTE" or state["waiting"] or state.get("stopped"):
        return None
    order = (read_json(session.dir / "strategy.json") or {}).get("order", [])
    left = [c for c in order if c not in state["results"]]
    if not left:
        return None
    counter = session.dir / "logs" / "stop_blocks.json"
    seen = read_json(counter, {"done": -1, "blocks": 0})
    blocks = seen["blocks"] if seen["done"] == len(state["results"]) else 0
    if blocks >= MAX_PUSHBACKS:
        write_json(counter, {"done": len(state["results"]), "blocks": blocks, "gave_up": True})
        return None
    write_json(counter, {"done": len(state["results"]), "blocks": blocks + 1})
    return (f"Session {session_id} is executing and {len(left)} case(s) have no result: {', '.join(left[:6])}. Keep going with `oqa run`, "
            f"or, if you are blocked, `oqa ask` the user or `oqa finish <case> --status BLOCKED --reason ...`. ({blocks + 1}/{MAX_PUSHBACKS})")


if __name__ == "__main__":
    try:
        json.load(sys.stdin)
        message = verdict()
    except Exception:
        sys.exit(0)
    if message:
        print(message, file=sys.stderr)
        sys.exit(2)
