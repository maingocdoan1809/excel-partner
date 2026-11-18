"""`oqa dashboard`: pick a session from the history and watch its cases and steps, with the picture of every step. Read-only, except the two things only
the user does: approve the strategy and answer the assistant's question."""
import json
import re
import unicodedata
from pathlib import Path
from urllib.parse import quote

from . import paths, report, sitemap, view, workspace
from .localserver import LocalApp, font_faces
from .state import GateError, Session
from .store import SchemaError

PAGE = Path(__file__).with_name("web") / "dashboard.html"
SITE_MAP_SCRIPT = PAGE.with_name("site_map.js")


def _attachment(name):
    """Content-Disposition for a download. A name with accents is sent twice: an ASCII one (the extension is kept) for old clients, and the real one as RFC 5987 UTF-8,
    so a browser never has to guess the bytes of a header and drop the `.xlsx` with them."""
    ascii_name = unicodedata.normalize("NFKD", name.replace("đ", "d").replace("Đ", "D")).encode("ascii", "ignore").decode("ascii")
    ascii_name = re.sub(r"[^\w.-]+", "_", ascii_name) or "file"
    if not ascii_name.lower().endswith(Path(name).suffix.lower()):
        ascii_name += Path(name).suffix
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(name)}"

class Dashboard(LocalApp):
    def __init__(self, select=None, timeout=4 * 3600):
        super().__init__(timeout)
        self.select = select            # (workspace, session) to open first
        self.quit = False

    def finished(self):
        return self.quit                # `oqa dashboard` of a newer skill asks an older dashboard to leave

    def page(self, query):
        boot = json.dumps({"select": list(self.select) if self.select else None}).replace("</", "<\\/")
        return PAGE.read_text(encoding="utf-8").replace("__TOKEN__", self.token).replace("__BOOT__", boot).replace("/*__FONT_FACES__*/", font_faces()).replace("/*__SITE_MAP_CSS__*/", SITE_MAP_SCRIPT.with_suffix(".css").read_text(encoding="utf-8")).replace("/*__SITE_MAP__*/", SITE_MAP_SCRIPT.read_text(encoding="utf-8"))

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
        return 200, target.read_bytes(), kind, {"Content-Disposition": _attachment(target.name)}

    def _sitemap(self, query):
        """Browse the saved workspace map. No browser, credentials, or Odoo writes."""
        ws = paths.resolve_workspace(query.get("ws"))
        config = workspace.load(ws)
        lang = sitemap.output_language(query.get("lang") or config.get("language"))
        if not sitemap.file_of(ws).is_file():
            return 200, {"workspace": ws, "available": False, "language": lang}
        data = sitemap.load(ws, lang)
        base = {"workspace": ws, "available": True, "language": lang, "meta": data["meta"], "header": sitemap.header(data, lang)}
        what = query.get("what", "show")
        if what == "show":
            display = sitemap.localized(data, lang)
            parents = {e["to"]: e["from"] for e in data["edges"] if e["kind"] == "menu"}
            opens = {e["from"]: e["to"] for e in data["edges"] if e["kind"] == "opens"}
            base["menus"] = [{"key": key, "label": node["label"], "name": node["label"].rsplit(" / ", 1)[-1],
                              "parent": parents.get(key), "action": opens.get(key), "xmlid": node.get("xmlid", "")}
                             for key, node in display["nodes"].items() if node["kind"] == "menu"]
            base["graph"] = {"nodes": [{"key": key, "kind": node["kind"], "type": node.get("type", "menu"),
                                         "label": node["label"].rsplit(" / ", 1)[-1] if node["kind"] == "menu" else node["label"],
                                         "trail": node["label"], "model": node.get("model", ""), "xmlid": node.get("xmlid", ""),
                                         "buttons": len(node.get("buttons", [])), "fields": len(node.get("fields", []))}
                                        for key, node in display["nodes"].items()],
                             "edges": [{k: v for k, v in edge.items() if k in ("from", "to", "kind", "label", "name")} for edge in display["edges"]]}
            base["unresolved"] = [{"where": u["where"], "why": sitemap.unresolved_text(u, lang), "language": u.get("language", "")} for u in data["unresolved"][:30]]
            base["unresolved_count"] = len(data["unresolved"])
            return 200, base
        if what == "find":
            kind = query.get("kind", "")
            if kind not in ("", "menu", "action", "view"):
                return 400, {"error": "kind must be menu, action or view"}
            text = query.get("q", "").strip()
            display = sitemap.localized(data, lang)
            found = sitemap.find(display, text, limit=60, kind=kind, lang=lang) if text else [(key, node, "") for key, node in display["nodes"].items() if not kind or node["kind"] == kind][:60]
            base["results"] = [{"key": key, "label": node["label"], "kind": node["kind"], "reason": reason,
                                "description": sitemap.describe(key, node, lang)}
                               for key, node, reason in found]
            return 200, base
        if what == "screen":
            key = query.get("screen", "")
            if key not in data["nodes"]:
                return 404, {"error": "No such screen in this workspace's sitemap."}
            display = sitemap.localized(data, lang)
            node = display["nodes"][key]
            view_type, match = query.get("view", ""), query.get("match", "")
            base.update(key=key, label=node["label"], description=sitemap.describe(key, node, lang),
                        views=sorted({data["nodes"][v]["type"] for v in sitemap.views_of(data, key)}),
                        buttons=sitemap.text_buttons(data, key, view=view_type, match=match, lang=lang),
                        fields=sitemap.text_fields(data, key, view=view_type, match=match, required=query.get("required") == "true", lang=lang),
                        path=sitemap.text_path(data, query.get("source", ""), key, lang=lang))
            view_keys = [v for v in sitemap.views_of(display, key) if not view_type or display["nodes"][v]["type"] == view_type]
            field_keys = view_keys
            if not view_type and any(display["nodes"][v]["type"] == "form" for v in field_keys):
                field_keys = [v for v in field_keys if display["nodes"][v]["type"] == "form"]

            def matches(item):
                return not match or all(word in sitemap._search_text(item) for word in sitemap._words(match))

            buttons, fields = [], []
            outgoing = sitemap._out(display)
            for view_key in view_keys:
                view_node = display["nodes"][view_key]
                targets = {e.get("name"): e["to"] for e in outgoing.get(view_key, []) if e["kind"] == "button"}
                for button in view_node.get("buttons", []):
                    if button["type"] in sitemap.OPEN_ROW or not matches(button):
                        continue
                    target = targets.get(button["name"])
                    buttons.append({**button, "view": view_node["type"], "selector": sitemap.selector_of("button", button["name"]),
                                    "target": target, "target_label": display["nodes"][target]["label"] if target else ""})
            for view_key in field_keys:
                view_node, seen = display["nodes"][view_key], set()
                for field in view_node.get("fields", []):
                    if field["name"] in seen or not matches(field) or (query.get("required") == "true" and not field["required"]):
                        continue
                    seen.add(field["name"])
                    fields.append({**field, "view": view_node["type"], "selector": sitemap.selector_of("field", field["name"], view_node["type"])})
            found_route = sitemap.path(display, query.get("source", ""), key)
            route = found_route or []
            route_keys = ([route[0]["from"]] + [edge["to"] for edge in route]) if route else ([key] if found_route is not None else [])
            base["controls"] = {"buttons": buttons[:60], "fields": fields[:60], "buttons_total": len(buttons), "fields_total": len(fields)}
            base["route"] = [{"key": route_key, "label": display["nodes"][route_key]["label"], "kind": display["nodes"][route_key]["kind"],
                              "via": route[index - 1]["kind"] if index else "start"} for index, route_key in enumerate(route_keys)]
            base["route_edges"] = route
            base["reachable"] = found_route is not None
            return 200, base
        return 400, {"error": "what must be show, find or screen"}

    def handle(self, method, path, query, body):
        try:
            if method == "GET" and path == "/api/sessions":
                return 200, {"sessions": view.list_sessions()}
            if method == "GET" and path == "/api/workspaces":
                return 200, {"workspaces": [{"name": name, "language": workspace.load(name).get("language", "en"), "has_sitemap": sitemap.file_of(name).is_file()}
                                             for name in paths.workspace_names()]}
            if method == "GET" and path == "/api/sitemap":
                return self._sitemap(query)
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
        except (GateError, SchemaError, paths.PathError, sitemap.SitemapError, json.JSONDecodeError, OSError) as error:
            return 200, {"error": "\n".join(error.problems) if isinstance(error, SchemaError) else str(error)}
        return 404, {"error": "not found"}
