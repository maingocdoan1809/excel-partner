"""workspace.json: the project's configuration (URL, profiles, reference folders). Never holds a password: a profile's secret lives in the
credential store under `secret_ref(workspace, profile)`."""
import re
from urllib.parse import urlparse

from . import SCHEMA_VERSION
from .paths import check_name, workspace_dir
from .secrets import SecretStore
from .store import SchemaError, read_json, write_json

LANGUAGES = ("vi", "ja", "en")
PREFIX_RE = re.compile(r"^[A-Za-z0-9_-]{2,30}$")


def secret_ref(ws, profile):
    return f"oqa:{ws}:{profile}"


def blank(name=""):
    return {"schema_version": SCHEMA_VERSION, "name": name, "url": "", "db": "", "profiles": [], "reference_dirs": [], "language": "vi", "test_prefix": "ZZ_AUTOTEST_"}


def normalize(raw):
    """What the form sends -> a clean workspace dict (no secrets). Unknown keys are dropped."""
    raw = raw if isinstance(raw, dict) else {}

    def text(key, default=""):
        return str(raw.get(key, default) or "").strip()

    url = text("url")
    if url and not re.match(r"^[a-z][a-z0-9+.-]*://", url, re.I):
        url = "http://" + url
    profiles = []
    for p in raw.get("profiles") or []:
        if isinstance(p, dict) and (str(p.get("name", "")).strip() or str(p.get("login", "")).strip()):
            profiles.append({"name": str(p.get("name", "")).strip(), "login": str(p.get("login", "")).strip(), "note": str(p.get("note", "")).strip()})
    return {"schema_version": SCHEMA_VERSION, "name": text("name"), "url": url.rstrip("/"), "db": text("db"), "profiles": profiles,
            "reference_dirs": [str(d).strip() for d in raw.get("reference_dirs") or [] if str(d).strip()],
            "language": text("language", "vi") or "vi", "test_prefix": text("test_prefix", "ZZ_AUTOTEST_") or "ZZ_AUTOTEST_"}


def validate(ws):
    problems = []
    try:
        check_name(ws.get("name", ""), "workspace name")
    except Exception as error:
        problems.append(str(error))
    parsed = urlparse(ws.get("url", ""))
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        problems.append("URL: enter the address you open in the browser, for example http://localhost:8069")
    if not ws.get("profiles"):
        problems.append("Profiles: add at least one (for example Admin).")
    seen = set()
    for p in ws.get("profiles", []):
        if not p.get("name") or not p.get("login"):
            problems.append("Each profile needs a name and a login.")
        elif p["name"].lower() in seen:
            problems.append(f"Profile '{p['name']}' appears twice.")
        else:
            try:
                check_name(p["name"], "profile name")
            except Exception as error:
                problems.append(str(error))
        seen.add(str(p.get("name", "")).lower())
    if ws.get("language") not in LANGUAGES:
        problems.append(f"Language must be one of {', '.join(LANGUAGES)}.")
    if not PREFIX_RE.match(ws.get("test_prefix", "")):
        problems.append("Test-data prefix: 2 to 30 letters, digits, - or _.")
    if problems:
        raise SchemaError(problems, "workspace")
    return ws


def load(name):
    ws = read_json(workspace_dir(name) / "workspace.json")
    if ws is None:
        raise SchemaError(f"workspace '{name}' does not exist", "workspace")
    return validate(ws)


def public_view(ws, store: SecretStore):
    """Safe to show anywhere (the form, `oqa ws show`, the assistant): each profile says only whether a password is saved."""
    view = dict(ws)
    view["profiles"] = [dict(p, has_secret=store.has(secret_ref(ws["name"], p["name"]))) for p in ws.get("profiles", [])]
    return view


def save(raw, store: SecretStore, passwords=None):
    """Validate, store the passwords given ({profile name: password}; missing or empty = keep the saved one), then write workspace.json.
    A profile that was removed loses its stored password."""
    ws = validate(normalize(raw))
    passwords = passwords or {}
    previous = read_json(workspace_dir(ws["name"]) / "workspace.json") or {}
    for profile in ws["profiles"]:
        ref = secret_ref(ws["name"], profile["name"])
        if not passwords.get(profile["name"]) and not store.has(ref):
            raise SchemaError(f"Profile '{profile['name']}' needs a password.", "workspace")
    for profile in ws["profiles"]:
        if passwords.get(profile["name"]):
            store.set(secret_ref(ws["name"], profile["name"]), passwords[profile["name"]])
    for old in previous.get("profiles", []):
        if old.get("name") not in {p["name"] for p in ws["profiles"]}:
            store.delete(secret_ref(ws["name"], old["name"]))
    write_json(workspace_dir(ws["name"]) / "workspace.json", ws)
    return ws
