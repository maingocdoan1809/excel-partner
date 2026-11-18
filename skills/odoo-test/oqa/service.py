"""The runner service: one background process per session that OWNS the browser, so a case that stopped at step 7 is repaired and picked up at step 7, on the page where
it stopped, instead of being run again from step 1 (which costs time, repeats side effects in Odoo and used to eat the repair budget before the last step was reached).

    oqa run  --client-->  runner.json (address + token)  --HTTP on 127.0.0.1-->  the service  --Playwright-->  the browser

The service is started by the first `oqa run` and leaves by itself when the session is no longer executing or nothing was asked for IDLE_SECONDS. It keeps nothing that
matters: the evidence and the state are files, written by the same Runner as without it. If it is gone, the next `oqa run` starts a new one and the case starts at step 1.
The browser is driven from ONE thread (Playwright's synchronous API belongs to the thread that made it); the HTTP threads only hand jobs over through a queue."""
import json
import os
import queue
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

from . import build_id
from .browser import launch
from .connect import LoginFailed
from .localserver import LocalApp
from .runtime import Runner
from .state import GateError, Session
from .store import SchemaError, read_json, write_json

IDLE_SECONDS = 45 * 60
START_SECONDS = 40
FILE = "runner.json"
REFUSED = (GateError, SchemaError, LoginFailed)


def pointer(session: Session) -> Path:
    return session.dir / FILE


class RunnerService(LocalApp):
    def __init__(self, session: Session, ws, store, headless=True, idle=IDLE_SECONDS):
        super().__init__(timeout=10 ** 9)
        self.session, self.ws, self.store, self.headless, self.idle = session, ws, store, headless, idle
        self.jobs, self.stopping, self.busy = queue.Queue(), threading.Event(), False

    def page(self, query):
        return ""

    def handle(self, method, path, query, body):
        if method == "GET" and path == "/ping":
            return 200, {"ok": True, "pid": os.getpid(), "headless": self.headless, "busy": self.busy}
        if method == "POST" and path in ("/run", "/explore", "/try"):
            reply = queue.Queue()
            self.jobs.put((path[1:], body, reply))
            return 200, reply.get()              # this request's thread waits; the browser thread answers
        if method == "POST" and path == "/stop":
            self.stopping.set()
            return 200, {"ok": True}
        return 404, {"error": "not found"}

    def _over(self):
        """The session no longer needs a browser: every case has a result, or it was closed."""
        try:
            return self.session.state()["phase"] != "EXECUTE"
        except Exception:
            return True

    def _run(self, runner, new_page, kind, body):
        try:
            if kind == "explore":
                return {"result": runner.explore(new_page, body["case"], to=body.get("to") or 1, needle=body.get("filter") or "")}
            if kind == "try":
                return {"result": runner.try_step(body["case"], body["step"], body.get("code") or "", body.get("filter") or "", bool(body.get("reset")), bool(body.get("undo")))}
            return {"result": runner.run(body["cases"], start=body.get("start"), fresh=bool(body.get("fresh")), new_page=new_page)}
        except REFUSED as error:
            return {"refused": "\n".join(error.problems) if isinstance(error, SchemaError) else str(error)}
        except Exception as error:
            return {"error": f"{type(error).__name__}: {str(error)[:400]}"}

    def serve(self, announce):
        """Run in the calling thread until asked to stop, idle for too long, or the session is over."""
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), self._handler())
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        announce({"url": f"http://127.0.0.1:{self.httpd.server_address[1]}", "token": self.token, "pid": os.getpid(), "headless": self.headless, "build": build_id()})
        last = time.monotonic()
        try:
            with launch(self.headless) as new_page:
                runner = Runner(self.session, self.ws, self.store, self.headless)
                while not self.stopping.is_set():
                    try:
                        kind, body, reply = self.jobs.get(timeout=1)
                    except queue.Empty:
                        if time.monotonic() - last > self.idle or self._over():
                            break
                        continue
                    self.busy = True
                    reply.put(self._run(runner, new_page, kind, body))
                    self.busy, last = False, time.monotonic()
                    if self._over():
                        break
        finally:
            mine = read_json(pointer(self.session), None)
            if mine and mine.get("pid") == os.getpid():
                pointer(self.session).unlink(missing_ok=True)
            self.httpd.shutdown()
            self.httpd.server_close()
            while not self.jobs.empty():
                self.jobs.get_nowait()[2].put({"error": "the runner stopped before it could run this"})


# --- the client: what `oqa run` uses -----------------------------------------------------------
def _call(info, method, path, body=None, timeout=60):
    request = urllib.request.Request(info["url"] + path, data=None if body is None else json.dumps(body).encode("utf-8"), method=method,
                                     headers={"X-OQA-Token": info["token"], "Content-Type": "application/json"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read() or b"{}")


def alive(session: Session):
    """The running service's address and ping, or None (a pointer left by a service that is gone is removed)."""
    info = read_json(pointer(session), None)
    if not info:
        return None
    try:
        return {**info, **_call(info, "GET", "/ping", timeout=3)}
    except Exception:
        pointer(session).unlink(missing_ok=True)
        return None


def stop(session: Session, wait=15):
    info = alive(session)
    if not info:
        return False
    try:
        _call(info, "POST", "/stop", {}, timeout=5)
    except Exception:
        pass
    deadline = time.monotonic() + wait
    while time.monotonic() < deadline and alive(session):
        time.sleep(0.3)
    return True


def spawn(session: Session, ws_name, headed):
    root = Path(__file__).resolve().parents[1]
    log = session.dir / "logs" / "runner.log"
    log.parent.mkdir(parents=True, exist_ok=True)
    command = [sys.executable, str(root / "oqa.py"), "runner", "serve", "--ws", ws_name, "--session", session.id] + (["--headed"] if headed else [])
    extra = ({"creationflags": subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {"start_new_session": True})
    with open(log, "ab") as out:
        subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=out, stderr=out, close_fds=True, **extra)
    deadline = time.monotonic() + START_SECONDS
    while time.monotonic() < deadline:
        info = alive(session)
        if info:
            return info
        time.sleep(0.4)
    raise GateError(f"The runner did not start within {START_SECONDS} seconds. Its log: {log}")


def ensure(session: Session, ws_name, headed):
    info = alive(session)
    if info and (info.get("headless") != (not headed) or info.get("build") != build_id()):         # a different window mode, or started from older files: start over
        stop(session)
        info = None
    return info or spawn(session, ws_name, headed)


def ask(session: Session, ws_name, kind, body, headed=False):
    """Hand the runner service (started if needed) a job: `run`, `explore` or `try`. Returns its result, or raises GateError with what was refused."""
    info = ensure(session, ws_name, headed)
    try:
        answer = _call(info, "POST", "/" + kind, body, timeout=3600)
    except (urllib.error.URLError, ConnectionError, OSError, ValueError) as error:
        pointer(session).unlink(missing_ok=True)
        raise GateError(f"The runner stopped while running ({str(error)[:120]}). Run again: it starts clean, at step 1 of the case. Its log: {session.dir / 'logs' / 'runner.log'}")
    if "refused" in answer:
        raise GateError(answer["refused"])
    if "error" in answer:
        raise GateError("The runner failed: " + answer["error"])
    return answer["result"]


def run(session: Session, ws_name, cases, start=None, fresh=False, headed=False):
    return ask(session, ws_name, "run", {"cases": cases, "start": start, "fresh": fresh}, headed)


def announce_to(session: Session):
    return lambda info: write_json(pointer(session), info)
