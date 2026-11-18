"""oqa: the command line of the odoo-test skill. Standard library only, except openpyxl (reading testcase workbooks) and Playwright (driving the browser)."""
VERSION = "0.1.0"
# Every JSON file oqa writes carries this number. A reader that meets a larger one must stop instead of guessing.
SCHEMA_VERSION = 1


def build_id():
    """A short digest of the skill's own files (code and pages). A long-lived process (the dashboard, the runner) records it when it starts: one that was started from older
    files is still answering with the code it loaded then, next to pages and commands that have since changed, so it must be replaced, not reused."""
    import hashlib
    from pathlib import Path
    here = Path(__file__).resolve().parent
    digest = hashlib.sha256()
    for file in sorted([*here.glob("*.py"), *here.glob("web/*.html"), *here.glob("web/*.js"), *here.glob("web/*.css")]):
        digest.update(file.name.encode("utf-8") + file.read_bytes())
    return digest.hexdigest()[:12]
