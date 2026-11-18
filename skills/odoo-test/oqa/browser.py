"""One place that starts Playwright: the same viewport and settings for connect, run and probe, so a picture taken anywhere looks the same."""
from contextlib import contextmanager

VIEWPORT = {"width": 1600, "height": 900}


@contextmanager
def launch(headless=True):
    """Yield a function `new_page(storage_state=None) -> (context, page)`. Everything is closed on exit."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        contexts = []

        def new_page(storage_state=None):
            context = browser.new_context(viewport=VIEWPORT, device_scale_factor=1, ignore_https_errors=True, storage_state=storage_state)
            contexts.append(context)
            return context, context.new_page()

        try:
            yield new_page
        finally:
            for context in contexts:
                try:
                    context.close()
                except Exception:
                    pass
            browser.close()
