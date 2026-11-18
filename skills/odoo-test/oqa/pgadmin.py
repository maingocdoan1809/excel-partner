"""Read data from the database THROUGH pgAdmin, so the evidence is a picture of a real application showing the real rows.

One fixed script, not something the assistant works out on each run: it opens pgAdmin, finds the server group / server / database in the tree, opens the Query
Tool, puts one SELECT in, runs it, reads the grid and photographs the Query Tool panel. Written against pgAdmin 9.x (Electron desktop, also the web build: the
page is the same). If pgAdmin shows something this script does not expect, it raises PgAdminError saying what it saw: nobody improvises selectors around it.

    cfg = {"group": "Servers", "server": "localhost", "database": "db_x", "login": "", "url": "", "exe": ""}
    with PgAdmin(cfg, password="") as pg:
        result = pg.query("select id, name from res_partner limit 5")   # {"columns", "rows", "total", "seconds"}
        pg.shoot("evidence.png")                                         # the Query Tool panel only: no server tree, no other servers' names

Desktop pgAdmin is driven over its DevTools port (started with --remote-debugging-port when it is not already open); with `url` set (pgAdmin in server mode) a
browser is opened on it and signed in with `login` / the password. The password is only typed into pgAdmin's own prompts and is never printed or kept."""
import os
import re
import subprocess
import time
import urllib.request
from pathlib import Path

PORT = 9222
SLOW = 60                     # seconds to wait for a server to connect, a query to answer
START = 180                   # a cold pgAdmin takes over a minute to start its server and replace the splash screen
# what changes data or structure, matched as a phrase so a column called `comment` or `lock` is not refused; even so, give pgAdmin a read-only database role
_WRITES = re.compile(r"\b(?:insert\s+into|delete\s+from|update\s+\S+\s+set|merge\s+into|call\s+\S+\s*\(|refresh\s+materialized|select\b[^;]*\binto\s+\S|(?:create|alter|drop|truncate)\s+(?:or\s+replace\s+)?(?:table|index|view|schema|database|extension|function|role|user|sequence|materialized)|truncate\s+\S|grant\s+\S+\s+on|revoke\s+\S+\s+on)", re.I)


class PgAdminError(Exception):
    pass


def check_select(sql):
    """The statement must be ONE read-only query. Returns it without trailing `;`, or raises."""
    text = re.sub(r"--[^\n]*|/\*.*?\*/", " ", sql or "", flags=re.S).strip().rstrip(";").strip()
    if not text:
        raise PgAdminError("the SQL is empty")
    if ";" in text:
        raise PgAdminError("one statement at a time: remove the ';' in the middle")
    if not re.match(r"(select|with|show|explain)\b", text, re.I):
        raise PgAdminError("only a SELECT (or WITH ... SELECT, SHOW, EXPLAIN) may be run")
    if _WRITES.search(re.sub(r"'(?:[^']|'')*'", "''", text)):                  # a word in a quoted string does not count
        raise PgAdminError("the query contains a statement that writes or changes something; only reading is allowed")
    return text


def check_expect(expect):
    """The rows a query must show: a non-empty list of {column: value}. Values are compared as the text pgAdmin shows, so they are returned as text."""
    if not isinstance(expect, list) or not expect:
        raise PgAdminError("expect must be a list of rows, each {column: value}")
    rows = []
    for row in expect:
        if not isinstance(row, dict) or not row:
            raise PgAdminError("each expected row is {column: value}, with at least one column")
        rows.append({str(k): str(v) for k, v in row.items()})
    return rows


def _exe(cfg):
    if cfg.get("exe"):
        return Path(cfg["exe"])
    roots = [os.environ.get("LOCALAPPDATA", "") + r"\Programs\pgAdmin 4", os.environ.get("ProgramFiles", "") + r"\pgAdmin 4"]
    for root in roots:
        found = sorted(Path(root).rglob("pgAdmin4.exe")) if root and Path(root).exists() else []
        if found:
            return found[0]
    raise PgAdminError("pgAdmin4.exe was not found: set 'exe' (its full path) or 'url' (pgAdmin in server mode)")


def _cdp_up():
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{PORT}/json/version", timeout=2).read()
        return True
    except Exception:
        return False


class PgAdmin:
    def __init__(self, cfg, password="", timeout=SLOW):
        for key in ("server", "database"):
            if not str(cfg.get(key, "")).strip():
                raise PgAdminError(f"pgAdmin settings: '{key}' is empty")
        self.cfg, self.password, self.timeout = cfg, password or "", timeout
        self._pw = self._browser = self.page = self.frame = None
        self._started = None

    # --- open / close ------------------------------------------------------------------
    def __enter__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        try:
            if self.cfg.get("url"):
                self._browser = self._pw.chromium.launch(headless=False, args=["--window-size=1600,1000"])
                self.page = self._browser.new_context(viewport={"width": 1600, "height": 1000}).new_page()
                self.page.goto(self.cfg["url"])
                self._sign_in()
            else:
                if not _cdp_up():
                    self._started = subprocess.Popen([str(_exe(self.cfg)), f"--remote-debugging-port={PORT}"])
                    self._wait(_cdp_up, "pgAdmin did not open its DevTools port")
                self._browser = self._pw.chromium.connect_over_cdp(f"http://127.0.0.1:{PORT}")
            self.page = self.page or self._main_page()
            self._wait(lambda: self.page.locator(".file-entry").count() > 0, "pgAdmin's tree did not appear")
        except BaseException:
            self.__exit__(None, None, None)
            raise
        return self

    def __exit__(self, *exc):
        try:
            if self.cfg.get("url") and self._browser:
                self._browser.close()                # a browser we opened; the desktop pgAdmin is left open for the person using it
        finally:
            if self._pw:
                self._pw.stop()
        return False

    def _wait(self, test, message, seconds=None):
        end = time.monotonic() + (seconds or self.timeout)
        while time.monotonic() < end:
            try:
                if test():
                    return
            except Exception:
                pass
            time.sleep(0.5)
        raise PgAdminError(message)

    def _main_page(self):
        pages = []

        def ready():
            pages[:] = [p for p in self._browser.contexts[0].pages if "/browser" in p.url]          # the splash page is replaced by /browser/ once the server is up
            return bool(pages)
        self._wait(ready, "pgAdmin is open but its main page did not load (still on the splash screen)", START)
        return pages[0]

    def _sign_in(self):
        p = self.page
        if p.locator("input[name=email], input[name=username]").count():
            if not self.cfg.get("login") or not self.password:
                raise PgAdminError("pgAdmin asks to sign in: give the login and the password")
            p.locator("input[name=email], input[name=username]").first.fill(self.cfg["login"])
            p.locator("input[name=password]").first.fill(self.password)
            p.keyboard.press("Enter")
            self._wait(lambda: "/browser" in p.url, "pgAdmin did not accept the login")

    # --- the tree ----------------------------------------------------------------------
    def _node(self, name):
        return self.page.locator(".file-entry").filter(has=self.page.locator(".file-name", has_text=re.compile(rf"^{re.escape(name)}(\s*\(\d+\))?$"))).first

    def _open(self, name, below):
        """Make `name` visible by opening `below` (once) when it is not."""
        if self._node(name).count():
            return
        self._node(below).dblclick()
        self._wait(lambda: self._node(name).count() > 0 or self._password_prompt(), f"'{name}' did not appear under '{below}'", self.timeout)
        if self._password_prompt():
            self._type_password()
            self._wait(lambda: self._node(name).count() > 0, f"'{name}' did not appear under '{below}' after the password was given")

    def _password_prompt(self):
        return self.page.locator("input[type=password]:visible").count() > 0

    def _type_password(self):
        if not self.password:
            raise PgAdminError(f"pgAdmin asks for the password of server '{self.cfg['server']}': give it (oqa dbcheck --ask-password)")
        self.page.locator("input[type=password]:visible").first.fill(self.password)
        self.page.keyboard.press("Enter")
        time.sleep(1.5)

    def _select_database(self):
        c = self.cfg
        chain = [n for n in (c.get("group"), c["server"], "Databases", c["database"]) if n]
        if not self._node(chain[0]).count():
            self._fail_missing(chain[0], chain)
        for parent, child in zip(chain, chain[1:]):
            if not self._node(parent).count():
                self._fail_missing(parent, chain)
            self._open(child, parent)
        self._node(c["database"]).click()

    def _fail_missing(self, name, chain):
        seen = [t.strip().replace("\n", " ") for t in self.page.locator(".file-entry").all_inner_texts()][:25]
        raise PgAdminError(f"'{name}' is not in pgAdmin's tree (looking for {' > '.join(chain)}). The tree shows: {seen}")

    # --- the Query Tool ----------------------------------------------------------------
    def _editor(self):
        self._select_database()
        old = {f.url for f in self.page.frames if "/sqleditor/" in f.url}               # tabs from an earlier use: the new one is the frame that was not there

        def up():
            self.frame = next((f for f in self.page.frames if "/sqleditor/" in f.url and f.url not in old), None)
            return self.frame is not None and self.frame.locator(".cm-content").count() > 0
        self.page.keyboard.press("Alt+Shift+Q")
        try:
            self._wait(up, "", 6)
        except PgAdminError:                                                     # not connected yet: the shortcut does nothing until the database is. Open the node (connects it), then ask again
            self._node(self.cfg["database"]).dblclick()
            self._wait(lambda: not self._connecting(), "pgAdmin is still connecting to the database", self.timeout)
            if self._password_prompt():
                self._type_password()
            self._node(self.cfg["database"]).click()
            self.page.keyboard.press("Alt+Shift+Q")
            try:
                self._wait(up, "", self.timeout)
            except PgAdminError:
                shot = Path(os.environ.get("TEMP", ".")) / "pgadmin_query_tool_failed.png"
                self.page.screenshot(path=str(shot))
                raise PgAdminError(f"the Query Tool did not open for database '{self.cfg['database']}' (what pgAdmin showed: {shot})")
        return self.frame.locator(".cm-content").first

    def _connecting(self):
        return self.page.locator("text=/Connecting to|Please wait/").count() > 0

    def query(self, sql):
        """Run one read-only query and read what the grid shows: {columns, rows (the ones on screen, at most the first page), total, seconds}."""
        sql = check_select(sql)
        editor = self._editor() if self.frame is None or self.frame.is_detached() else self.frame.locator(".cm-content").first
        editor.click()
        self.page.keyboard.press("Control+A")
        self.page.keyboard.press("Delete")
        self.page.keyboard.insert_text(sql)
        started = time.monotonic()
        before = self._status()
        self.page.keyboard.press("F5")
        self._wait(lambda: self._status() not in (before, "", None) or self._error(), "the query did not finish", self.timeout)
        if self._error():
            raise PgAdminError("the database refused the query: " + self._error())
        time.sleep(0.8)                                                          # the grid paints a moment after the footer
        grid = self.frame.evaluate("""() => { const g = document.querySelector('[role=grid]');
            const cols = [...g.querySelectorAll('[role=columnheader]')].slice(1).map(e => e.innerText.split('\\n')[0].trim());
            const rows = [...g.querySelectorAll('[role=row]')].map(r => [...r.querySelectorAll('[role=gridcell]')].slice(1).map(c => c.innerText)).filter(r => r.length);
            return {cols, rows}; }""")
        match = re.search(r"Total rows:\s*(\d+)", self._status() or "")
        return {"columns": grid["cols"], "rows": grid["rows"], "total": int(match.group(1)) if match else len(grid["rows"]), "seconds": round(time.monotonic() - started, 1)}

    def _status(self):
        loc = self.frame.locator("text=/Total rows:/")
        return loc.first.inner_text() + " " + (self.frame.locator("text=/Query complete/").first.inner_text() if self.frame.locator("text=/Query complete/").count() else "") if loc.count() else ""

    def _error(self):
        loc = self.frame.locator("text=/^ERROR:/")
        return loc.first.inner_text()[:300] if loc.count() else ""

    def shoot(self, path):
        """Photograph the Query Tool panel (the SQL and its result). The object tree, with the other servers' and databases' names, is not in the picture."""
        box = self.page.locator("iframe").first.bounding_box()
        if not box:
            raise PgAdminError("the Query Tool panel is not on screen")
        self.page.screenshot(path=str(path), clip={"x": box["x"], "y": box["y"], "width": box["width"], "height": box["height"] - 2})
        return Path(path)
