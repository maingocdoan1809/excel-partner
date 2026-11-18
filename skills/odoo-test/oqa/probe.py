"""`oqa probe`: what can be clicked and filled on a page, in a few lines. The cheap way for the assistant to find the selector that fits THIS Odoo
(instead of guessing, or reading the whole page).

Two sources: a page opened now (`probe`), and the page where a step stopped, saved at that moment (`failed`), which needs no browser and is the screen the
assistant actually has to fix: opening the landing page shows nothing of a form six steps in."""
import re

from . import browser, connect
from .odoo_ui import Odoo

JS = """
([needle, limit]) => {
  const vis = e => { const b = e.getBoundingClientRect(), s = getComputedStyle(e); return b.width > 1 && b.height > 1 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const out = [];
  for (const e of document.querySelectorAll('button, a[href], input:not([type=hidden]), select, textarea, [role=button], [role=menuitem], [role=tab], .o_field_widget, th[data-name], .o_data_row:first-child')) {
    if (!vis(e)) continue;
    const text = (e.innerText || e.value || e.placeholder || e.getAttribute('aria-label') || '').replace(/\\s+/g, ' ').trim().slice(0, 40);
    const own = e.getAttribute('name') || (e.closest('[name]') && e.closest('[name]').getAttribute('name')) || e.getAttribute('data-name') || '';
    const classes = [...e.classList].filter(c => /^o_|^btn|^nav|^dropdown/.test(c)).slice(0, 2).join('.');
    const line = e.tagName.toLowerCase() + (e.id ? '#' + e.id : '') + (classes ? '.' + classes : '') + (own ? '[name=' + own + ']' : '') + (e.getAttribute('data-menu-xmlid') ? '[menu=' + e.getAttribute('data-menu-xmlid') + ']' : '') + (text ? ' "' + text + '"' : '');
    if (!needle || line.toLowerCase().includes(needle.toLowerCase())) out.push(line);
  }
  return [...new Set(out)].slice(0, limit);
}
"""
SHOWN = 70                    # lines `oqa probe` prints
KEPT = 400                    # lines saved from the page where a step stopped
FAILED_FILE = "failed_page.txt"


def probe(session, ws, store, profile_name=None, path=None, needle="", headless=True):
    profile = next((p for p in ws["profiles"] if p["name"] == profile_name), ws["profiles"][0])
    with browser.launch(headless) as new_page:
        _, page = new_page(connect.storage_path(session, profile["name"]) if session else None)
        known = ((session.state().get("connect") or {}).get("profiles", {}).get(profile["name"]) or {}).get("landing_url") if session else None
        landing = path or known
        page.goto(landing if landing and landing.startswith("http") else ws["url"] + (landing or "/web"), wait_until="domcontentloaded")
        if "/web/login" in page.url:
            connect.login(page, ws, profile, store)
        Odoo(page, ws["url"]).ready()
        return {"url": page.url, "title": page.title(), "elements": page.evaluate(JS, [needle, SHOWN])}


def save_page(page, file):
    """Write the page as it is now (address, title, every control `probe` knows) to `file`: called when a step fails, while the page is still where it failed."""
    elements = page.evaluate(JS, ["", KEPT])
    file.parent.mkdir(parents=True, exist_ok=True)
    file.write_text("\n".join([f"url\t{page.url}", f"title\t{page.title()}", *elements]) + "\n", encoding="utf-8")


def _read(file):
    lines = file.read_text(encoding="utf-8").splitlines()
    head = {key: value for key, _, value in (line.partition("\t") for line in lines[:2])}
    return head.get("url", ""), head.get("title", ""), lines[2:]


def failed(session, case=None, needle=""):
    """The page where the latest failed step of `case` (or of any case) stopped, filtered by `needle`: {url, title, elements, file}. No browser needed."""
    folder = session.dir / "evidence"
    name = re.sub(r"[^\w.-]+", "_", case) if case else "*"
    pattern = f"{name}/a*/{FAILED_FILE}"
    saved = sorted(folder.glob(pattern), key=lambda f: f.stat().st_mtime)
    if not saved:
        raise FileNotFoundError("no failed step has saved its page yet" + (f" for {case}" if case else "") + ": run the case first")
    url, title, lines = _read(saved[-1])
    shown = [line for line in lines if not needle or needle.lower() in line.lower()]
    return {"url": url, "title": title, "elements": shown[:SHOWN], "total": len(lines), "file": str(saved[-1])}


QUOTED = re.compile(r"[\"']([^\"'\n]{2,40})[\"']")
CLASSES = re.compile(r"\.(o_[\w-]+)")
NAMES = re.compile(r"name=['\"]?(\w+)")


def near(lines, reason, per_word=3, limit=8):
    """The few controls of the page that look like what the failed step was after: any text, class or field name quoted in the failure reason."""
    words = [w for w in QUOTED.findall(reason) + CLASSES.findall(reason) + NAMES.findall(reason) if not re.fullmatch(r"[\d\s.]+(ms)?", w) and "Timeout" not in w]
    out = []
    for word in dict.fromkeys(words):
        hits = [line for line in lines if word.lower() in line.lower() and line not in out]
        out += hits[:per_word]
    return out[:limit]
