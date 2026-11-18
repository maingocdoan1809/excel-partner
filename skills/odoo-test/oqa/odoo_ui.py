"""Odoo's web client, as one-line actions for case scripts. Written for Odoo 16-18 class names; when a selector does not fit THIS Odoo, the case
script (or knowledge/index.md) says what does, found with `oqa probe`. Every method waits for the client to settle and raises with a message that
says what was not found, so a failed step reads as a sentence, not a stack trace."""
import re
import time

from . import diagnose

TIMEOUT = 8000
READ_METHODS = ("search_read", "search_count", "read", "read_group", "fields_get", "name_search", "check_access_rights", "get_views")      # what `o.rpc` will call: reading only
MAX_ROWS = 200
QUIET_MS = 400
BUSY = ".o_loading_indicator:not(.d-none), .o_blockUI, .o_loading:not(.d-none), .fa-spin, .o_view_loading, .placeholder-glow"
BUSY_FREE = "() => ![...document.querySelectorAll(%r)].some(e => { const b = e.getBoundingClientRect(); return b.width > 0 && b.height > 0; })" % BUSY
SIGNATURE = """() => { const root = document.querySelector('.o_action_manager') || document.body;
  const busy = [...document.querySelectorAll(%r)].filter(e => { const b = e.getBoundingClientRect(); return b.width > 0 && b.height > 0; }).length;
  return [root.querySelectorAll('button, input, select, textarea, .o_field_widget, .o_data_row').length, root.innerText.length, busy]; }""" % BUSY


class Odoo:
    def __init__(self, page, base_url, prefix="ZZ_AUTOTEST_", on_check=None, timeout=None):
        self.page, self.base, self.prefix = page, base_url.rstrip("/"), prefix
        self.on_check = on_check or (lambda: None)           # told each time a check is made, so the runner can tell a step that checked nothing
        self.timeout = timeout or TIMEOUT                    # ms; a slow page gets more through the case's timeout_s, never by a sleep in the script
        page.set_default_timeout(self.timeout)

    # --- waiting and moving around -----------------------------------------------------
    def ready(self):
        """The web client is up, not busy, and what it shows has stopped changing (records loaded, widgets rendered). Used by every helper and by `oqa try`,
        so exploring and running agree on when a page is "done". It never raises on a page that keeps moving: it returns after at most the case's timeout."""
        p = self.page
        p.wait_for_selector(".o_web_client", state="visible")
        p.wait_for_function(BUSY_FREE)
        try:
            p.wait_for_load_state("networkidle", timeout=3000)       # the menus and the first view arrive after the shell; long polling may keep it from ever idling
        except Exception:
            pass
        self._quiet()
        return self

    def _quiet(self):
        """Wait until the screen's shape (controls, text length, spinners) is the same for QUIET_MS in a row, or the case's timeout passes."""
        p, deadline = self.page, time.monotonic() + self.timeout / 1000
        last, since = None, time.monotonic()
        while time.monotonic() < deadline:
            try:
                now = p.evaluate(SIGNATURE)
            except Exception:                                        # the page navigated under us: look again
                last = None
                p.wait_for_timeout(100)
                continue
            if now != last:
                last, since = now, time.monotonic()
            elif now[2] == 0 and (time.monotonic() - since) * 1000 >= QUIET_MS:
                return
            p.wait_for_timeout(100)

    def goto(self, path):
        self.page.goto(path if path.startswith("http") else self.base + path)
        return self.ready()

    def menu(self, *names):
        """menu("Sales", "Orders"): the app, then each menu below it."""
        p = self.page
        sections = p.locator(".o_menu_sections")
        if not sections.get_by_text(names[0], exact=True).count() and not p.locator(".o_menu_brand", has_text=names[0]).count():
            p.locator(".o_navbar_apps_menu button, .o_navbar_apps_menu > a, .o_menu_toggle").first.click()
            p.locator(".dropdown-menu:visible").get_by_text(names[0], exact=True).first.click()
            self.ready()
        for name in names[1:]:
            dropdown = p.locator(".dropdown-menu:visible").get_by_text(name, exact=True)
            (dropdown if dropdown.count() else sections.get_by_text(name, exact=True)).first.click()
            self.ready()
        return self

    # --- finding things --------------------------------------------------------------------
    def button(self, text):
        """A visible button in the current view or dialog, by its text."""
        scope = self.page.locator(".modal:visible, .o_control_panel, .o_form_view, .o_content").first
        found = scope.get_by_role("button", name=text, exact=True)
        return (found if found.count() else self.page.get_by_role("button", name=text, exact=True)).first

    def field(self, name):
        """The input of a form field, by its technical name (the `name=` of the widget)."""
        return self.page.locator(f".o_field_widget[name='{name}'] input, .o_field_widget[name='{name}'] textarea, .o_field_widget[name='{name}'] select").first

    def row(self, text):
        return self.page.locator(".o_data_row", has_text=text).first

    # --- looking things up (read only) ---------------------------------------------------
    def rpc(self, model, method, args=None, kwargs=None):
        """One READ-ONLY call to Odoo's ORM with the signed-in session (the profile's own rights apply). Writing methods are refused here, whatever the script says."""
        if method not in READ_METHODS:
            raise PermissionError(f"o.rpc only reads ({', '.join(READ_METHODS)}); '{method}' changes data: do it through the screen")
        body = {"jsonrpc": "2.0", "method": "call", "params": {"model": model, "method": method, "args": args or [], "kwargs": kwargs or {}}}
        answer = self.page.request.post(f"{self.base}/web/dataset/call_kw/{model}/{method}", data=body).json()
        if answer.get("error"):
            error = answer["error"]
            raise AssertionError(f"{model}.{method}: {(error.get('data') or {}).get('message') or error.get('message')}")
        return answer.get("result")

    def records(self, model, domain=None, fields=None, limit=20, order=None):
        """search_read: [{field: value}] (a many2one is [id, name]). At most MAX_ROWS."""
        kwargs = {"fields": fields or [], "limit": min(int(limit), MAX_ROWS)}
        if order:
            kwargs["order"] = order
        return self.rpc(model, "search_read", [domain or []], kwargs)

    def models(self, like="", limit=30):
        """[(technical name, label)] of the models whose name or label contains `like`: is it `sale.order` or `sale.orders`?"""
        domain = ["|", ["model", "ilike", like], ["name", "ilike", like]] if like else []
        return [(r["model"], r["name"]) for r in self.records("ir.model", domain, ["model", "name"], limit, "model")]

    def model_exists(self, model):
        return bool(self.rpc("ir.model", "search_count", [[["model", "=", model]]]))

    def fields(self, model, like="", limit=60):
        """[{name, string, type, relation, required, readonly}] of a model's fields whose name or label contains `like`: the technical name behind a label on screen."""
        info = self.rpc(model, "fields_get", [], {"attributes": ["string", "type", "relation", "required", "readonly"]})
        needle = like.lower()
        rows = [{"name": n, **{k: v for k, v in a.items() if v not in (None, False, "")}} for n, a in sorted(info.items())
                if not needle or needle in n.lower() or needle in str(a.get("string", "")).lower()]
        return rows[:limit]

    def xmlid(self, ref):
        """{model, res_id, ref} for an external id like `base.main_company` or `sale.group_sale_manager`, or None when it does not exist."""
        module, _, name = ref.partition(".")
        found = self.records("ir.model.data", [["module", "=", module], ["name", "=", name]], ["model", "res_id"], 1)
        return {"ref": ref, "model": found[0]["model"], "res_id": found[0]["res_id"]} if found else None

    def find(self, model, text, limit=5):
        """[(id, display name)] of the records of `model` whose name contains `text` (name_search): is there such a partner / product / user?"""
        return [tuple(r) for r in self.rpc(model, "name_search", [], {"name": text, "limit": min(int(limit), MAX_ROWS)})]

    def rows(self):
        return self.page.locator(".o_data_row").count()

    def pager(self):
        """(first, last, total) of the list's pager, or (0, rows, rows) when the list has none."""
        p = self.page
        if not p.locator(".o_pager_value").count():
            return 0, self.rows(), self.rows()
        value = p.locator(".o_pager_value").first.inner_text().replace("–", "-")
        total = int(re.sub(r"\D", "", p.locator(".o_pager_limit").first.inner_text()) or 0)
        low, _, high = value.partition("-")
        return int(low or 0), int(high or low or 0), total

    def toasts(self):
        return [t.strip() for t in self.page.locator(".o_notification").all_inner_texts() if t.strip()]

    # --- doing things -----------------------------------------------------------------------
    def new(self):
        self.page.locator(".o_list_button_add, .o_form_button_create, .o-kanban-button-new").first.click()
        self.page.wait_for_selector(".o_form_view")
        return self.ready()

    def fill(self, name, value):
        self.field(name).fill(str(value))
        return self

    def m2o(self, name, text):
        """Pick a record in a many2one: type, then click the suggestion that contains `text`."""
        box = self.field(name)
        box.fill(text)
        self.page.locator(".o-autocomplete--dropdown-item, .ui-autocomplete .ui-menu-item, .o_m2o_dropdown_option").filter(has_text=text).first.click()
        try:                    # Odoo 17+ puts the chosen record into the input a moment AFTER the click: wait for it, or the next read sees an empty field
            self.page.wait_for_function("([n, t]) => { const i = document.querySelector(`.o_field_widget[name='${n}'] input`); return !!i && i.value.toLowerCase().includes(t.toLowerCase()); }",
                                        arg=[name, text], timeout=self.timeout)
        except Exception:
            raise AssertionError(f"'{text}' was clicked in {name} but never appeared in the field")
        return self

    def save(self):
        """Save the form; raise with what the page says if it did not save."""
        self.page.locator(".o_form_button_save").first.click()
        try:
            self.page.wait_for_selector(".o_form_button_save", state="hidden", timeout=self.timeout * 3 // 8)
        except Exception:
            raise AssertionError("the form did not save: " + diagnose.read(self.page)["summary"])
        return self.ready()

    def upload(self, target, fixture):
        """Attach a prepared file (`s.fixture(name)`, or a path) to a file field. `target` is the `<input type=file>` or the button / widget that opens the file chooser
        (a selector or a locator). Waits for the client to settle, so the next line sees the uploaded file."""
        path = str(getattr(fixture, "path", fixture))
        found = self.page.locator(target) if isinstance(target, str) else target
        if found.first.evaluate("e => e.tagName === 'INPUT' && e.type === 'file'"):
            found.first.set_input_files(path)
        else:
            with self.page.expect_file_chooser() as chooser:
                found.first.click()
            chooser.value.set_files(path)
        return self.ready()

    def search(self, text):
        box = self.page.locator(".o_searchview_input").first
        box.fill(text)
        box.press("Enter")
        return self.ready()

    def open_row(self, text):
        self.row(text).click()
        self.page.wait_for_selector(".o_form_view")
        return self.ready()

    # --- checking ---------------------------------------------------------------------------
    def expect_text(self, text):
        self.on_check()
        self.page.get_by_text(text, exact=False).first.wait_for(state="visible")

    def expect_visible(self, selector):
        self.on_check()
        self.page.locator(selector).first.wait_for(state="visible")

    def value(self, name):
        """What the field shows: the text of an input, the chosen label of a select, or the text of a read-only widget (a badge, a statusbar field)."""
        widget = self.page.locator(f".o_field_widget[name='{name}']").first
        box = widget.locator("input:not([type=hidden]), textarea, select").first
        if box.count():
            if box.evaluate("e => e.tagName") == "SELECT":
                return box.evaluate("e => e.selectedOptions[0] ? e.selectedOptions[0].text : ''").strip()
            return box.input_value()
        return widget.inner_text().strip()

    def expect_value(self, name, value):
        """The field shows `value`. It keeps looking until the client has settled (an onchange can change the value after a click), then fails with what it saw."""
        self.on_check()
        deadline = time.monotonic() + self.timeout / 1000
        while True:
            actual = self.value(name)
            if actual == str(value):
                return
            if time.monotonic() > deadline:
                raise AssertionError(f"field {name} is '{actual}', expected '{value}'")
            self.page.wait_for_timeout(150)

    def expect_rows(self, count):
        self.on_check()
        actual = self.rows()
        assert actual == count, f"the list has {actual} rows, expected {count}"

    def expect_no_errors(self):
        self.on_check()
        d = diagnose.read(self.page)
        assert not d["errors"] and not d["fields"], d["summary"]

    def reach(self, enough, make, limit=200):
        """Make test data until `enough()` is true (more than a page of rows, a tag that overflows its cell): the answer to 'the environment has too little data'.
        `make(i)` creates one record whose name starts with the test prefix. Gives up after `limit` records."""
        made = 0
        while not enough():
            assert made < limit, f"still not enough after creating {limit} records"
            make(made)
            made += 1
        return made

    def unique(self, label):
        """Like name(), plus a time suffix: for models that refuse a duplicate name, so a second run never meets the first run's record."""
        return f"{self.prefix}{label}-{time.strftime('%H%M%S')}"

    def name(self, label):
        """A record name that starts with the test prefix, so cleanup finds exactly what the test made."""
        return f"{self.prefix}{label}"
