"""JSON files on disk: written atomically, versioned, and guarded by a lock where two writers could meet."""
import json
import os
import time
from contextlib import contextmanager
from pathlib import Path

from . import SCHEMA_VERSION


class SchemaError(Exception):
    """A file or a submission that does not match its schema. `problems` lists every defect, so the caller fixes them in one go."""

    def __init__(self, problems, what="data"):
        self.problems = [problems] if isinstance(problems, str) else list(problems)
        self.what = what
        super().__init__(f"{what}: " + "; ".join(self.problems))


def read_json(path, default=None):
    """The parsed file; `default` when it does not exist. A newer schema_version than this code knows is an error, not a guess."""
    path = Path(path)
    if not path.exists():
        return default
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except ValueError as error:
        raise SchemaError(f"{path.name} is not valid JSON ({error})", path.name)
    version = data.get("schema_version") if isinstance(data, dict) else None
    if isinstance(version, int) and version > SCHEMA_VERSION:
        raise SchemaError(f"{path.name} was written by a newer oqa (schema {version}, this one reads {SCHEMA_VERSION}). Update the skill.", path.name)
    return data


def write_json(path, data):
    """Write to a temp file and rename, so a reader never sees half a file."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, path)


@contextmanager
def locked(path, timeout=10.0):
    """One writer at a time for read-modify-write of a file. A lock older than a minute is taken to be abandoned."""
    lock = Path(str(path) + ".lock")
    lock.parent.mkdir(parents=True, exist_ok=True)
    deadline = time.monotonic() + timeout
    while True:
        try:
            os.close(os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
            break
        except FileExistsError:
            try:
                if time.time() - lock.stat().st_mtime > 60:
                    lock.unlink(missing_ok=True)
                    continue
            except OSError:
                pass
            if time.monotonic() > deadline:
                raise TimeoutError(f"{lock.name} is held by another oqa process.")
            time.sleep(0.05)
    try:
        yield
    finally:
        lock.unlink(missing_ok=True)
