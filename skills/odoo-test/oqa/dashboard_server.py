"""`oqa dashboard`: pick a session from the history and watch its cases and steps, with the picture of every step. Read-only, except the two things only
the user does: approve the strategy and answer the assistant's question."""
import json
from pathlib import Path

from . import paths, report, view, workspace
from .localserver import LocalApp, font_faces
from .state import GateError, Session
from .store import SchemaError

PAGE = Path(__file__).with_name("web") / "dashboard.html"


class Dashboard(LocalApp):
    def __init__(self, select=None, timeout=4 * 3600):
        super().__init__(timeout)
        self.select = select            # (workspace, session) to open first
        self.quit = False

    def finished(self):
        return self.quit                # `oqa dashboard` of a newer skill asks an older dashboard to leave

    def page(self, query):
        boot = json.dumps({"select": list(self.select) if self.select else None}).replace("</", "<\\/")
        return PAGE.read_text(encoding="utf-8").replace("__TOKEN__", self.token).replace("__BOOT__", boot).replace("/*__FONT_FACES__*/", font_faces())

    def _session(self, params):
        return Session(paths.resolve_workspace(params.get("ws")), paths.check_name(params.get("id", ""), "session id"))

    def _image(self, query):
        session = self._session(query)
        root = (session.dir / "evidence").resolve()
        target = (session.dir / query.get("path", "")).resolve()
        if root not in target.parents or target.suffix.lower() not in (".png", ".jpg") or not target.is_file():
            return 404, {"error": "no such picture"}
        return 200, target.read_bytes(), "image/png" if target.suffix.lower() == ".png" else "image/jpeg"

    def _report_file(self, query):
        session = self._session(query)
        root = (session.dir / "report").resolve()
        target = (root / query.get("name", "")).resolve()
        if root not in target.parents or target.suffix.lower() not in (".xlsx", ".md") or not target.is_file():
            return 404, {"error": "no such file"}
        kind = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if target.suffix.lower() == ".xlsx" else "text/markdown"
        return 200, target.read_bytes(), kind, {"Content-Disposition": f'attachment; filename="{target.name}"'}

    def handle(self, method, path, query, body):
        try:
            if method == "GET" and path == "/api/sessions":
                return 200, {"sessions": view.list_sessions()}
            if method == "GET" and path == "/api/session":
                return 200, view.session_view(self._session(query))
            if method == "GET" and path == "/api/image":
                return self._image(query)
            if method == "POST" and path == "/api/confirm":
                self._session(body).confirm_current("dashboard")
                return 200, {"ok": True}
            if method == "GET" and path == "/api/report":
                return self._report_file(query)
            if method == "POST" and path == "/api/report":
                session = self._session(body)
                built = report.build(session, workspace.load(session.ws))
                return 200, {"ok": True, "main": Path(built["main"]).name}
            if method == "POST" and path == "/api/quit":
                self.quit = True
                return 200, {"ok": True}
            if method == "POST" and path == "/api/stop":
                self._session(body).stop("dashboard")
                return 200, {"ok": True}
            if method == "POST" and path == "/api/resume":
                self._session(body).resume("dashboard")
                return 200, {"ok": True}
            if method == "POST" and path == "/api/hint":
                self._session(body).hint(str(body.get("case", "")), str(body.get("text", "")), "dashboard")
                return 200, {"ok": True}
            if method == "POST" and path == "/api/answer":
                if not str(body.get("text", "")).strip():
                    return 200, {"error": "Write the answer first."}
                self._session(body).answer(str(body["text"]).strip())
                return 200, {"ok": True}
        except (GateError, SchemaError, paths.PathError) as error:
            return 200, {"error": "\n".join(error.problems) if isinstance(error, SchemaError) else str(error)}
        return 404, {"error": "not found"}
