"""CONNECT: sign in once per profile, keep the signed-in state for the run, and record what this Odoo does after login (the landing page, the version).
Odoo is not like other apps here: the page after login differs by version (/web, /odoo), a database selector can come first, and two-factor
stops a script. Each of those is reported as a sentence that says what to do, not as a timeout."""
from pathlib import Path

from . import browser, knowledge
from .secrets import SecretStore
from .state import GateError, Session
from .workspace import secret_ref

SETTLED = "() => document.querySelector('.o_web_client') || document.querySelector('.alert-danger') || document.querySelector('input[name=totp_token]')"
VERSION_JS = ("() => fetch('/web/webclient/version_info', {method: 'POST', headers: {'Content-Type': 'application/json'}, "
              "body: JSON.stringify({jsonrpc: '2.0', method: 'call', params: {}})}).then(r => r.json()).then(j => (j.result || {}).server_version || '').catch(() => '')")


SLOW_MS = 90000          # the first load after an Odoo restart builds its asset bundles, which can take a minute


class LoginFailed(Exception):
    pass


def login(page, ws, profile, store: SecretStore):
    """Sign `profile` in on `page`. Returns {landing_url, title, version}; raises LoginFailed with the reason."""
    password = store.get(secret_ref(ws["name"], profile["name"]))
    if password is None:
        raise LoginFailed(f"no password is saved for profile {profile['name']}: run oqa init and enter it")
    page.goto(ws["url"] + "/web/login" + (f"?db={ws['db']}" if ws.get("db") else ""), wait_until="domcontentloaded")
    if "/web/database/selector" in page.url:                 # not `.o_database_list`: the real login page has a "Manage Databases" link with that class
        raise LoginFailed("Odoo asks which database to use: enter the database name in oqa init")
    page.fill("input[name=login]", profile["login"])
    page.fill("input[name=password]", password)
    page.click("button[type=submit]")
    page.wait_for_function(SETTLED, timeout=SLOW_MS)
    if page.locator("input[name=totp_token]").count():
        raise LoginFailed(f"profile {profile['name']} asks for a two-factor code: only the user can enter it (use --headed), or turn it off for the test account")
    if page.locator(".alert-danger").count():
        raise LoginFailed(f"Odoo refused the sign-in: {page.locator('.alert-danger').first.inner_text().strip()[:160]}")
    page.wait_for_selector(".o_web_client", state="visible", timeout=SLOW_MS)
    return {"landing_url": page.url, "title": page.title(), "version": page.evaluate(VERSION_JS) or ""}


def connect(session: Session, ws, store: SecretStore, headless=True):
    """Sign every profile in, store each signed-in state under the session (auth/, never read by the assistant), then move CONNECT -> CONFIRM."""
    session.require("CONNECT")
    auth = session.dir / "auth"
    auth.mkdir(exist_ok=True)
    info, failures = {}, []
    with browser.launch(headless) as new_page:
        for profile in ws["profiles"]:
            context, page = new_page()
            try:
                info[profile["name"]] = login(page, ws, profile, store)
                context.storage_state(path=str(auth / f"{profile['name']}.json"))
            except Exception as error:
                picture = session.dir / "logs" / f"connect_{profile['name']}.png"
                picture.parent.mkdir(exist_ok=True)
                try:
                    page.screenshot(path=str(picture))
                except Exception:
                    picture = None
                failures.append(f"{profile['name']}: {str(error).splitlines()[0][:200]}" + (f" (picture: {picture})" if picture else ""))
    if failures:
        raise GateError("Could not sign in:\n  " + "\n  ".join(failures) + "\nFix the cause (oqa init for a wrong login or database), then run oqa connect again.")
    for name, item in info.items():
        path = "/" + item["landing_url"].split("//", 1)[-1].partition("/")[2]
        knowledge.learn(ws["name"], "login", "login", f"After signing in as {name} the page is {path.split('?')[0]}" + (f" (Odoo {item['version']})" if item["version"] else ""))
    session.connected({"profiles": info, "version": next(iter(info.values()))["version"]})
    return info


def storage_path(session: Session, profile):
    path = session.dir / "auth" / f"{profile}.json"
    return str(path) if path.exists() else None
