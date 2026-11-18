"""A short-lived web page served to the person at this computer: `oqa init` (the form) and `oqa dashboard` (progress) are both one of these.

It listens on 127.0.0.1 only, on a random port. Every request needs the one-time token (header or `t=` in the address, so an <img> can carry it),
and Host / Origin must be this server, so no other page open in the browser can call it."""
import hmac
import json
import secrets
import threading
import time
import webbrowser
from pathlib import Path
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

MAX_BODY = 64 * 1024
FONTS = Path(__file__).with_name("web") / "fonts"


def font_faces():
    """@font-face rules for Be Vietnam Pro. The computer's own copy is used when there is one (local()); files dropped into oqa/web/fonts are added only when they exist,
    so a page never asks for a font file that is not there (that would be a 404 in the console)."""
    out = []
    for weight, name in ((300, "Light"), (400, "Regular"), (500, "Medium"), (600, "SemiBold")):
        sources = [f'local("BeVietnamPro-{name}")', f'local("Be Vietnam Pro {name}")']
        for ext, kind in (("woff2", "woff2"), ("ttf", "truetype")):
            if (FONTS / f"BeVietnamPro-{name}.{ext}").is_file():
                sources.append(f'url("/font/BeVietnamPro-{name}.{ext}") format("{kind}")')
        out.append(f'@font-face {{ font-family:"Be Vietnam Pro"; font-weight:{weight}; font-style:normal; font-display:swap; src:{", ".join(sources)}; }}')
    return " ".join(out)


class LocalApp:
    """Subclasses give `page()` (the HTML) and `handle(method, path, query, body)` -> (status, payload[, content type]); payload is a dict (JSON) or bytes."""

    def __init__(self, timeout=900):
        self.token = secrets.token_urlsafe(16)
        self.timeout = timeout
        self.httpd = None

    def page(self, query):
        raise NotImplementedError

    def handle(self, method, path, query, body):
        return 404, {"error": "not found"}

    def finished(self):
        """True when the server should stop (the form was saved)."""
        return False

    def _handler(self):
        app = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def _send(self, code, payload, kind=None, headers=None):
                if isinstance(payload, bytes):
                    data, kind = payload, kind or "application/octet-stream"
                elif isinstance(payload, str):
                    data, kind = payload.encode("utf-8"), kind or "text/html"
                else:
                    data, kind = json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json"
                self.send_response(code)
                self.send_header("Content-Type", kind + ("; charset=utf-8" if kind.startswith(("text/", "application/json")) else ""))
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                for key, value in (headers or {}).items():
                    self.send_header(key, value)
                self.end_headers()
                self.wfile.write(data)

            def _same_site(self):
                port = self.server.server_address[1]
                origin = self.headers.get("Origin")
                return self.headers.get("Host", "") in (f"127.0.0.1:{port}", f"localhost:{port}") and (not origin or origin in (f"http://127.0.0.1:{port}", f"http://localhost:{port}"))

            def _allowed(self, query):
                port = self.server.server_address[1]
                if self.headers.get("Host", "") not in (f"127.0.0.1:{port}", f"localhost:{port}"):
                    return False
                origin = self.headers.get("Origin")
                if origin and origin not in (f"http://127.0.0.1:{port}", f"http://localhost:{port}"):
                    return False
                given = self.headers.get("X-OQA-Token") or (query.get("t") or [""])[0]
                return hmac.compare_digest(given.encode(), app.token.encode())

            def _font(self, name):
                """Fonts the person dropped into oqa/web/fonts (Be Vietnam Pro, for example). A stylesheet cannot carry the token, so this one route needs none: it serves only
                font files from that folder, to a page of this same server."""
                target = (FONTS / name).resolve()
                if FONTS.resolve() not in target.parents or target.suffix.lower() not in (".woff2", ".ttf", ".otf") or not target.is_file():
                    return self._send(404, {"error": "no such font"})
                self._send(200, target.read_bytes(), "font/woff2" if target.suffix.lower() == ".woff2" else "font/ttf")

            def _serve(self, method):
                url = urlparse(self.path)
                query = parse_qs(url.query)
                if method == "GET" and url.path.startswith("/font/") and self._same_site():
                    return self._font(url.path[len("/font/"):])
                if not self._allowed(query):
                    return self._send(403, {"error": "forbidden"})
                body = {}
                if method == "POST":
                    length = int(self.headers.get("Content-Length") or 0)
                    if length > MAX_BODY:
                        return self._send(413, {"error": "too large"})
                    try:
                        body = json.loads(self.rfile.read(length) or b"{}")
                    except ValueError:
                        return self._send(400, {"error": "bad json"})
                if url.path == "/" and method == "GET":
                    return self._send(200, app.page({k: v[0] for k, v in query.items()}), "text/html")
                try:
                    self._send(*app.handle(method, url.path, {k: v[0] for k, v in query.items()}, body))
                except Exception as error:
                    self._send(500, {"error": str(error)})

            def do_GET(self):
                self._serve("GET")

            def do_POST(self):
                self._serve("POST")

        return Handler

    def serve(self, announce, open_browser=True, forever=False):
        """Serve until `finished()` (or, with `forever`, until the time is up or the user presses Ctrl+C). The address, token included, goes to `announce`."""
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), self._handler())
        self.httpd.timeout = 0.4
        url = f"http://127.0.0.1:{self.httpd.server_address[1]}/?t={self.token}"
        announce(url)
        if open_browser:
            threading.Thread(target=lambda: webbrowser.open(url), daemon=True).start()
        deadline = time.monotonic() + self.timeout
        try:
            while not self.finished() and time.monotonic() < deadline:
                self.httpd.handle_request()
            time.sleep(0.3 if self.finished() else 0)          # let the browser receive the reply before the socket goes away
        except KeyboardInterrupt:
            pass
        finally:
            self.httpd.server_close()
