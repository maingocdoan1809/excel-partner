# -*- coding: utf-8 -*-
"""Preflight (`/system-test check`, stage 1 of `run`). Read-only. Prints a checklist and writes data/preflight.json.
Checks: config file, app URL reachable, browser starts / attaches, login state, project data check (hooks.check_data), snapshot source,
role / environment switching configured for the matrix, catalog present and up to date, suites load."""
import json, sys, urllib.request
from ..core import browser, hooks, session
from ..core.config import cfg, WORK

sys.stdout.reconfigure(encoding="utf-8")


def main():
    rows = []

    def add(item, ok, detail, level="ERROR"):
        rows.append(dict(item=item, ok=bool(ok), detail=str(detail), level="OK" if ok else level))

    add("config", cfg.path.exists(), str(cfg.path) if cfg.path.exists() else f"missing {cfg.path}: run /autotest-init")
    add("source", True, "source.use=true: the agent may read the app source" if cfg.use_source else "source.use=false: black-box, the agent does not read the app source")
    try:
        r = urllib.request.urlopen(cfg.base_url, timeout=10)
        add("app URL", True, f"{cfg.base_url} -> HTTP {r.status}")
    except Exception as e:
        code = getattr(e, "code", None)
        add("app URL", code is not None and code < 500, f"{cfg.base_url} -> {code or e}", "ERROR")
    try:
        pg = browser.page()
        pg.goto(cfg.url(cfg.raw["app"].get("home_path", "/")), wait_until="domcontentloaded")
        pg.wait_for_timeout(1500)
        add("browser", True, f"mode={cfg.browser.get('mode')} channel={cfg.browser.get('channel')} url={pg.url}")
        li = session.logged_in(pg)
        if not li and cfg.raw["auth"].get("type") == "form":
            try:
                session.login(pg)
                li = session.logged_in(pg)
            except Exception as e:
                add("login", False, f"form login failed: {e}")
        add("login", li, f"auth.type={cfg.raw['auth'].get('type')} logged_in={li}" + ("" if li else " -> log in once in the test browser (manual) or fix auth settings"),
            "PENDING")
    except Exception as e:
        add("browser", False, f"cannot start / attach: {e}")
    finally:
        browser.stop()
    fn = getattr(hooks.module(), "check_data", None)
    if fn:
        try:
            msg = fn()
            if isinstance(msg, dict):                                  # structured result: {"status": VERIFIED|VIOLATION|UNVERIFIED|ERROR, "detail": "..."}
                status, detail = str(msg.get("status", "UNVERIFIED")).upper(), msg.get("detail", "")
            else:
                text = str(msg)
                status = "ERROR" if text.startswith("ERROR") else "UNVERIFIED" if text.lower().startswith(("not implemented", "unverified")) else "VERIFIED"
                detail = text
            add("data source", status == "VERIFIED", f"{status}: {detail}", "ERROR" if status in ("ERROR", "VIOLATION") else "WARN")
        except Exception as e:
            add("data source", False, f"hooks.check_data failed: {e}")
    else:
        add("data source", False, "no hooks.check_data(): the framework cannot prove it reads the database the app uses", "WARN")
    add("data snapshot", cfg.raw["data"].get("snapshot", "none") != "none", f"data.snapshot={cfg.raw['data'].get('snapshot')}"
        + (" -> integrity of the data will NOT be verified" if cfg.raw["data"].get("snapshot", "none") == "none" else ""), "WARN")
    M = cfg.raw["matrix"]
    if len(M.get("roles", {})) > 1:
        add("role switching", cfg.raw["roles"].get("switch", "none") != "none", f"roles.switch={cfg.raw['roles'].get('switch')} for roles {list(M['roles'])}", "WARN")
    if len(M.get("envs", {})) > 1:
        add("env switching", cfg.raw["envs"].get("switch", "none") != "none", f"envs.switch={cfg.raw['envs'].get('switch')} for envs {list(M['envs'])}", "WARN")
    try:
        from . import catalog
        cat = catalog.load()
        if not cat:
            add("catalog", False, "no catalog yet: /system-test update", "WARN")
        else:
            st = catalog.stale(cat)
            add("catalog", not st, f"build {cat['build']}, {len(cat['cases'])} IDs" + (f"; out of date: {st}" if st else ""), "WARN")
            if cat.get("suite_errors"):
                add("suites", False, cat["suite_errors"])
    except Exception as e:
        add("catalog", False, e, "WARN")
    for r in rows:
        print(f"[{r['level']:7}] {r['item']:15} {r['detail']}")
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    (cfg.data_dir / "preflight.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if any(r["level"] == "ERROR" for r in rows) else 0


if __name__ == "__main__":
    sys.exit(main())
