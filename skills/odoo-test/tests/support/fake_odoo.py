"""A tiny stand-in for Odoo's web client: the login page, the page after login, an app menu, a list with a pager, a form with a required field and a
many2one, a save button that validates, notifications, and the version endpoint. It uses Odoo's own class names so oqa's helpers meet familiar markup.
It is NOT Odoo: it proves the runner, the evidence and the repair loop work; selectors on a real Odoo are checked with `oqa probe`."""
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

USERS = {"admin": "pw-admin", "lead": "pw-lead"}
PAGE_SIZE = 5
LOGIN = """<!doctype html><html><head><title>Login</title></head><body><div class="oe_login_form">
%ERROR%<form method="post" action="/web/login"><input name="login" placeholder="Email"><input name="password" type="password"><button type="submit" class="btn btn-primary">Log in</button></form><a class="o_database_list" href="/web/database/selector">Manage Databases</a></div></body></html>"""
APP = """<!doctype html><html><head><title>Odoo</title><style>
body{font:14px sans-serif;margin:0}.o_navbar{background:#714b67;color:#fff;padding:8px;display:flex;gap:16px;align-items:center}
.dropdown-menu{display:none;position:absolute;background:#fff;color:#000;border:1px solid #ccc;min-width:160px;z-index:9}.dropdown-menu.show{display:block}
.dropdown-item{display:block;padding:6px 10px;cursor:pointer}.o_notification{position:fixed;right:16px;top:60px;background:#fff;border:1px solid #888;padding:10px;z-index:10}
.o_notification.border-danger{border-color:#c00;color:#c00}.o_field_invalid input{border:2px solid #c00}.o_control_panel{padding:10px;display:flex;gap:10px;align-items:center}
.o_content{padding:10px}table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #ddd;padding:6px;text-align:left}.o_field_widget{margin:8px 0;display:block}
.d-none{display:none}</style></head><body><div class="o_web_client">
<nav class="o_navbar"><div class="o_navbar_apps_menu"><button class="dropdown-toggle" id="apps">Apps</button><div class="dropdown-menu" id="appsmenu"><a class="dropdown-item o_app" id="app-sales">Sales</a></div></div>
<span class="o_menu_brand"></span><div class="o_menu_sections"></div></nav><div class="o_action_manager" id="main"></div><div id="toasts"></div></div>
<script>
let orders = [], customers = ["Azure Interior", "Deco Addict"], page = 0, dirty = false;
const $ = s => document.querySelector(s);
const toast = (text, bad) => { const n = document.createElement('div'); n.className = 'o_notification' + (bad ? ' border-danger' : ''); n.textContent = text; $('#toasts').append(n); setTimeout(() => n.remove(), 4000); };
const load = async () => { orders = await (await fetch('/api/orders')).json(); };
$('#apps').onclick = () => $('#appsmenu').classList.toggle('show');
$('#app-sales').onclick = () => { $('#appsmenu').classList.remove('show'); $('.o_menu_brand').textContent = 'Sales';
  $('.o_menu_sections').innerHTML = '<a class="o_nav_entry" id="nav-orders">Orders</a>'; $('#nav-orders').onclick = showList; };
async function showList() { await load(); const total = orders.length, rows = orders.slice(page * %PAGE%, page * %PAGE% + %PAGE%);
  $('#main').innerHTML = '<div class="o_control_panel"><button class="btn o_list_button_add">New</button><input class="o_searchview_input" placeholder="Search..."><span class="o_pager"><span class="o_pager_value">' + (total ? (page * %PAGE% + 1) + '-' + (page * %PAGE% + rows.length) : '0') + '</span> / <span class="o_pager_limit">' + total + '</span> <button class="o_pager_next">&gt;</button></span></div><div class="o_content o_list_view"><table><thead><tr><th>Name</th><th>Customer</th></tr></thead><tbody>' + rows.map(o => '<tr class="o_data_row"><td>' + o.name + '</td><td>' + o.partner + '</td></tr>').join('') + '</tbody></table></div>';
  $('.o_list_button_add').onclick = showForm; $('.o_pager_next').onclick = () => { if ((page + 1) * %PAGE% < total) { page++; showList(); } };
  document.querySelectorAll('.o_data_row').forEach(r => r.onclick = () => showForm(r.children[0].textContent)); }
function showForm() { dirty = false;
  $('#main').innerHTML = '<div class="o_control_panel"><button class="btn o_form_button_save d-none">Save</button></div><div class="o_content o_form_view"><div class="o_field_widget o_required_modifier" name="name"><label>Name</label> <input></div><div class="o_field_widget" name="partner_id"><label>Customer</label> <input autocomplete="off"><div class="dropdown-menu o-autocomplete--dropdown-menu"></div></div></div>';
  const save = $('.o_form_button_save'), name = $('[name=name] input'), partner = $('[name=partner_id] input'), menu = $('.o-autocomplete--dropdown-menu');
  const touch = () => { dirty = true; save.classList.remove('d-none'); };
  name.oninput = touch; partner.oninput = () => { touch(); menu.innerHTML = customers.filter(c => c.toLowerCase().includes(partner.value.toLowerCase())).map(c => '<a class="dropdown-item o-autocomplete--dropdown-item">' + c + '</a>').join(''); menu.classList.add('show');
    menu.querySelectorAll('a').forEach(a => a.onclick = () => { menu.classList.remove('show'); const v = a.textContent; setTimeout(() => { partner.value = v; }, 400); }); };
  save.onclick = async () => { if (!name.value.trim()) { $('[name=name]').classList.add('o_field_invalid'); toast('Invalid fields: Name', true); return; }
    await fetch('/api/orders', {method: 'POST', body: JSON.stringify({name: name.value, partner: partner.value || '-'})}); toast('Saved'); save.classList.add('d-none'); setTimeout(showList, 200); }; }
</script></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _send(self, code, body, kind="text/html", headers=None):
        data = body.encode() if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", kind + "; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)

    def _user(self):
        cookie = self.headers.get("Cookie", "")
        return cookie.split("session=")[1].split(";")[0] if "session=" in cookie else None

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/web/login":
            return self._send(200, LOGIN.replace("%ERROR%", ""))
        if path in ("/odoo", "/web"):
            return self._send(200, APP.replace("%PAGE%", str(PAGE_SIZE))) if self._user() in USERS else self._send(303, "", headers={"Location": "/web/login"})
        if path == "/api/orders":
            return self._send(200, json.dumps(self.server.orders), "application/json")
        self._send(404, "not found")

    def do_POST(self):
        body = self.rfile.read(int(self.headers.get("Content-Length") or 0)).decode()
        if self.path == "/web/login":
            form = dict(p.split("=", 1) for p in body.split("&") if "=" in p)
            from urllib.parse import unquote_plus
            login, password = unquote_plus(form.get("login", "")), unquote_plus(form.get("password", ""))
            if USERS.get(login) == password:
                return self._send(303, "", headers={"Location": "/odoo", "Set-Cookie": f"session={login}; Path=/"})
            return self._send(200, LOGIN.replace("%ERROR%", '<div class="alert alert-danger">Wrong login/password</div>'))
        if self.path == "/web/webclient/version_info":
            return self._send(200, json.dumps({"result": {"server_version": "17.0"}}), "application/json")
        if self.path == "/api/orders":
            self.server.orders.append(json.loads(body))
            return self._send(200, "{}", "application/json")
        self._send(404, "not found")


def start(orders=3):
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    server.orders = [{"name": f"S0000{i}", "partner": "Azure Interior"} for i in range(1, orders + 1)]
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server, f"http://127.0.0.1:{server.server_address[1]}"
