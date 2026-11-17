# -*- coding: utf-8 -*-
"""Remove the test objects (name starts with the test prefix) that a run created.
  python -m qakit.tools.cleanup            # plan only (read-only): what would be removed
  python -m qakit.tools.cleanup --apply    # remove, through the project's CLEANERS (project hooks), dependencies first
Sources: safety/created_registry.json of the run (guard.track) + hooks.discover() (leftovers of earlier runs that still exist).
Objects without a cleaner for their kind are listed as MANUAL. Never touches anything without the prefix."""
import sys
from ..core import guard, hooks
from ..core.config import cfg

sys.stdout.reconfigure(encoding="utf-8")


def plan():
    H = hooks.module()
    items = {(r["kind"], r["name"]): r for r in guard.tracked()}
    disc = getattr(H, "discover", None)
    if disc:
        for r in disc() or []:
            items.setdefault((r["kind"], r["name"]), r)
    items = [r for r in items.values() if guard.is_test(r["name"])]
    order = list(getattr(H, "CLEANUP_ORDER", []))
    return sorted(items, key=lambda r: (order.index(r["kind"]) if r["kind"] in order else len(order), r["name"]))


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    apply = "--apply" in argv
    H = hooks.module()
    cleaners = getattr(H, "CLEANERS", {})
    items = plan()
    if not items:
        print("nothing to clean (no test object found)")
        return 0
    manual = failed = 0
    for r in items:
        fn = cleaners.get(r["kind"])
        if not fn:
            manual += 1
            print(f"MANUAL   {r['kind']:15} {r['name']}  (no cleaner for this kind in project hooks)")
            continue
        if not apply:
            print(f"PLAN     {r['kind']:15} {r['name']}")
            continue
        try:
            ok = fn(r)
            print(f"{'REMOVED' if ok else 'FAILED':8} {r['kind']:15} {r['name']}")
            if ok:
                guard.untrack(r["kind"], r["name"])
            else:
                failed += 1
        except Exception as e:
            failed += 1
            print(f"FAILED   {r['kind']:15} {r['name']}: {e}")
    print(f"{len(items)} test object(s); manual: {manual}" + ("" if apply else " (plan only: add --apply to remove)"))
    try:
        from ..core import browser
        browser.stop()
    except Exception:
        pass
    return 1 if (failed or (apply and manual)) else 0


if __name__ == "__main__":
    sys.exit(main())
