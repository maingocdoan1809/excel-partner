# -*- coding: utf-8 -*-
"""Generic UI primitives (any web UI). Screen-specific code lives in the project's pages/ (page objects).
Rules baked in:
 * prefer selectors by role / label / test id / structure; visible text only through label tuples (one entry per UI language),
   e.g. SAVE = ("Save", "保存", "Lưu") -> button(pg, SAVE);
 * the caller declares the risk of an action (core.guard) - these primitives only perform it;
 * waits are explicit (toasts disappear after a few seconds, dialogs initialise asynchronously)."""
import re
from . import session

TOAST = "[role=status], [role=alert], [data-sonner-toast], .toast, .Toastify__toast"
DIALOG = "[role=dialog], [role=alertdialog], dialog[open]"


def goto(pg, url, settle=800):
    """Open a URL (absolute, or a path of the app), wait for it to settle, log in again when the session expired."""
    from .config import cfg
    url = url if "://" in url else cfg.url(url)
    pg.goto(url)
    pg.wait_for_load_state("networkidle")
    pg.wait_for_timeout(settle)
    if session.ensure_session(pg):
        pg.goto(url)
        pg.wait_for_load_state("networkidle")
        pg.wait_for_timeout(settle)


def _rx(names):
    return re.compile("^\\s*(" + "|".join(re.escape(n) for n in names) + ")\\s*$")


def button(scope, names, exact=True):
    """Button by any of its labels (tuple, one per UI language) inside `scope` (page or locator)."""
    names = (names,) if isinstance(names, str) else names
    return scope.get_by_role("button", name=_rx(names) if exact else re.compile("|".join(map(re.escape, names))))


def click_button(scope, names, timeout=10000):
    b = button(scope, names).first
    b.wait_for(timeout=timeout)
    b.click()


def top_dialog(pg):
    return pg.locator(DIALOG).last


def dialog_open(pg):
    return pg.locator(DIALOG).count() > 0


def close_dialog(pg):
    if dialog_open(pg):
        pg.keyboard.press("Escape")
        pg.wait_for_timeout(500)


def toast(pg, timeout=6000):
    """Text of the first toast / status message ('' if none appeared)."""
    loc = pg.locator(TOAST).first
    try:
        loc.wait_for(timeout=timeout)
        return loc.inner_text().replace("\n", " ")
    except Exception:
        return ""


def table_row(pg, text):
    """First table row containing `text`."""
    return pg.locator("tr", has_text=text).first


def stable_text(pg, selector, polls=2, interval=2000, max_polls=45):
    """Wait until the text of `selector` stops changing (streaming answers, progress). Returns the final text."""
    prev, stable = None, 0
    for _ in range(max_polls):
        pg.wait_for_timeout(interval)
        cnt = pg.locator(selector).count()
        txt = pg.locator(selector).last.inner_text() if cnt else ""
        stable = stable + 1 if (cnt and txt == prev) else 0
        prev = txt
        if stable >= polls:
            break
    return prev or ""


def errors(pg):
    """Visible validation / error texts (role=alert, aria-invalid descriptions, common error classes)."""
    return pg.evaluate("""()=>[...document.querySelectorAll('[role=alert],[aria-live=assertive],.error,.invalid-feedback,[class*=error-message],[class*=text-red],[class*=text-destructive]')]
        .filter(e=>e.offsetParent).map(e=>e.innerText.trim()).filter(Boolean)""")
