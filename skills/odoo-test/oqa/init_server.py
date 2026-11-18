"""`oqa init`: the configuration form. One HTML file shipped inside oqa (nothing is copied anywhere), served by a LocalApp that stops when the form is
saved or when the time runs out."""
import json
import os
import ssl
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

from . import paths, workspace
from .localserver import LocalApp, font_faces
from .secrets import SecretStore
from .store import SchemaError

PAGE = Path(__file__).with_name("web") / "init.html"


def check_url(url, timeout=6):
    """Is the address reachable, and does it look like Odoo? No sign-in is attempted here. A certificate that does not verify is reported, not fatal."""
    if urlparse(url).scheme not in ("http", "https"):
        return {"ok": False, "error": "Address must start with http:// or https://"}
    last = None
    for verify in (True, False):
        context = None if verify else ssl._create_unverified_context()
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "oqa-check"})
            with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
                body = response.read(20000).decode("utf-8", "replace").lower()
                final = response.geturl()
                return {"ok": True, "status": response.status, "final_url": final, "tls_unverified": not verify,
                        "odoo": "odoo" in body or "/web/login" in final or "o_login" in body}
        except urllib.error.HTTPError as error:
            return {"ok": True, "status": error.code, "final_url": url, "odoo": False, "tls_unverified": not verify}
        except Exception as error:
            last = error
            reason = getattr(error, "reason", error)
            if verify and url.startswith("https") and isinstance(reason, ssl.SSLError):
                continue                      # a certificate that does not verify: try again without verifying, and say so
            break
    return {"ok": False, "error": f"Không mở được {url}: {getattr(last, 'reason', last)}"}


class InitServer(LocalApp):
    def __init__(self, store: SecretStore, name=None, timeout=900):
        super().__init__(timeout)
        self.store, self.name, self.result = store, name, None

    def finished(self):
        return self.result is not None

    def _boot(self):
        names = paths.workspace_names()
        chosen = self.name if self.name in names else (names[0] if len(names) == 1 and not self.name else None)
        return {"names": names, "home": str(paths.home()), "homeFixed": bool(os.environ.get(paths.HOME_ENV)), "suggest": self.name or "",
                "workspace": workspace.public_view(workspace.load(chosen), self.store) if chosen else None}

    def page(self, query):
        boot = json.dumps(self._boot(), ensure_ascii=False).replace("</", "<\\/")
        return PAGE.read_text(encoding="utf-8").replace("__TOKEN__", self.token).replace("__BOOT__", boot).replace("/*__FONT_FACES__*/", font_faces())

    def save(self, body):
        ws = body.get("workspace") if isinstance(body.get("workspace"), dict) else {}
        home = str(body.get("home") or "").strip()
        if home and Path(home).expanduser() != paths.home():
            if os.environ.get(paths.HOME_ENV):
                raise SchemaError(f"The storage folder is fixed by the {paths.HOME_ENV} variable on this computer.", "workspace")
            paths.set_home(home)
        saved = workspace.save(ws, self.store, body.get("passwords") if isinstance(body.get("passwords"), dict) else {})
        self.result = {"workspace": saved["name"], "home": str(paths.home()), "url": saved["url"], "profiles": [p["name"] for p in saved["profiles"]],
                       "reference_dirs": saved["reference_dirs"]}
        return {"ok": True, "name": saved["name"]}

    def handle(self, method, path, query, body):
        try:
            if method == "GET" and path == "/api/workspace":
                name = query.get("name", "")
                return 200, {"workspace": workspace.public_view(workspace.load(name), self.store) if name else workspace.blank()}
            if method == "POST" and path == "/api/check":
                return 200, check_url(str(body.get("url", "")).strip())
            if method == "POST" and path == "/api/save":
                return 200, self.save(body)
        except SchemaError as error:
            return 200, {"error": "\n".join(error.problems)}
        except (paths.PathError, ValueError) as error:
            return 200, {"error": str(error)}
        except Exception as error:      # e.g. the credential store refused
            return 200, {"error": f"Không lưu được: {error}"}
        return 404, {"error": "not found"}

    def run(self, announce, open_browser=True):
        """Serve until saved or timed out. Returns the saved summary, or None on timeout."""
        self.serve(announce, open_browser)
        return self.result


def status(name=None):
    """Is a workspace configured and valid? (no server, no secrets)"""
    names = paths.workspace_names()
    pick = name or (names[0] if len(names) == 1 else None)
    if not pick:
        return {"configured": False, "workspaces": names}
    try:
        ws = workspace.load(pick)
    except (SchemaError, paths.PathError) as error:
        return {"configured": False, "workspaces": names, "problem": str(error)}
    return {"configured": True, "workspace": ws["name"], "url": ws["url"], "profiles": [p["name"] for p in ws["profiles"]], "reference_dirs": ws["reference_dirs"], "workspaces": names}
