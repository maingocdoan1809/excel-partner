"""`oqa doctor`: is this computer ready to run the skill, and `--fix` installs what is missing (pip packages and the Playwright browser)."""
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

from . import VERSION
from .secrets import SecretStoreUnavailable, default_store

MIN_PYTHON = (3, 9)
PACKAGES = {"openpyxl": "openpyxl", "playwright": "playwright", "PIL": "pillow"}


def _browser_installed():
    base = Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH") or (Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ms-playwright"))
    return base.is_dir() and any(p.name.startswith("chromium") for p in base.iterdir())


def check():
    missing = [pip for module, pip in PACKAGES.items() if importlib.util.find_spec(module) is None]
    try:
        default_store()
        secrets_ok, secrets_note = True, ""
    except SecretStoreUnavailable as error:
        secrets_ok, secrets_note = False, str(error)
    python_ok = sys.version_info[:2] >= MIN_PYTHON
    browser = "playwright" not in missing and _browser_installed()
    return {"oqa": VERSION, "python": sys.version.split()[0], "python_ok": python_ok, "missing_packages": missing, "browser_installed": browser,
            "secret_store": secrets_ok, "secret_store_note": secrets_note, "ready": python_ok and not missing and browser and secrets_ok}


def fix(report, log):
    """Install the missing packages, then the browser. `log(line)` receives progress; returns the new report."""
    steps = []
    if report["missing_packages"]:
        steps.append([sys.executable, "-m", "pip", "install", *report["missing_packages"]])
    if "playwright" in report["missing_packages"] or not report["browser_installed"]:
        steps.append([sys.executable, "-m", "playwright", "install", "chromium"])
    for step in steps:
        log("$ " + " ".join(step))
        subprocess.run(step, check=True)
    return check()
