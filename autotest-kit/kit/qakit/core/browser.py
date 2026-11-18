# -*- coding: utf-8 -*-
"""Browser for the test run (Playwright). Two modes (qa.config.json "browser.mode"):
  launch : the framework starts its own browser with a persistent profile (browser.profile_dir) - logins survive between runs;
           channel = chromium | msedge | chrome; headless true/false.
  cdp    : attach to a browser the user started with --remote-debugging-port=<cdp_port> --user-data-dir=<profile> (useful when the
           login needs MFA / SSO that only a human can pass once).
Screenshots are taken from the TEST PAGE ONLY (page.screenshot, viewport): never the screen, never another window or application."""
import io
from pathlib import Path
from PIL import Image
from playwright.sync_api import sync_playwright
from .config import cfg, WORK

_pw = _ctx = _browser = _page = None


def _start():
    global _pw, _ctx, _browser
    b = cfg.browser
    _pw = sync_playwright().start()
    if b.get("mode") == "cdp":
        _browser = _pw.chromium.connect_over_cdp(f"http://127.0.0.1:{cfg.cdp_port}")
        _ctx = _browser.contexts[0] if _browser.contexts else _browser.new_context()
    elif cfg.raw["auth"].get("type") == "session":
        st = cfg.raw["auth"].get("storage_state") or ""
        if not st or not Path(st).exists():
            raise RuntimeError(f"the saved sign-in '{cfg.raw['auth'].get('session_name', '')}' was not found: sign in again at Settings > Testing > Sign-ins")
        w, h = b.get("viewport", [1600, 900])
        kw = dict(headless=bool(b.get("headless")), slow_mo=b.get("slow_mo", 0) or 0)
        if b.get("channel") and b.get("channel") != "chromium":
            kw["channel"] = b["channel"]
        _browser = _pw.chromium.launch(**kw)
        _ctx = _browser.new_context(storage_state=st, viewport={"width": w, "height": h})      # read only: the saved state is never written back
    else:
        w, h = b.get("viewport", [1600, 900])
        kw = dict(headless=bool(b.get("headless")), viewport={"width": w, "height": h}, slow_mo=b.get("slow_mo", 0) or 0)
        ch = b.get("channel")
        if ch and ch != "chromium":
            kw["channel"] = ch
        prof = WORK / b.get("profile_dir", "browser_profile")
        prof.mkdir(parents=True, exist_ok=True)
        _ctx = _pw.chromium.launch_persistent_context(str(prof), **kw)


def page():
    """The test page (created on first use)."""
    global _page
    if _pw is None:
        _start()
    if _page is None or _page.is_closed():
        _page = _ctx.pages[0] if _ctx.pages else _ctx.new_page()
        _page.set_default_timeout(int(cfg.browser.get("action_timeout_ms", 8000)))     # Playwright's own default (30 s) makes every failed attempt take half a minute
    return _page


def new_tab():
    page()
    return _ctx.new_page()


def stop():
    """Release the browser. In cdp mode the user's browser stays open; in launch mode it is closed."""
    global _pw, _ctx, _browser, _page
    if _pw:
        try:
            if cfg.browser.get("mode") != "cdp" and _ctx:
                _ctx.close()
                if _browser:
                    _browser.close()
        finally:
            _pw.stop()
    _pw = _ctx = _browser = _page = None


def grab(pg):
    """(PIL image of the visible viewport in CSS pixels, viewport size). Box coordinates from bounding_box() map 1:1 onto it."""
    png = pg.screenshot(scale="css", animations="disabled", caret="initial")
    img = Image.open(io.BytesIO(png)).convert("RGB")
    return img
