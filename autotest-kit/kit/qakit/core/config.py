# -*- coding: utf-8 -*-
"""Central configuration of one project: <autotest root>/qa.config.json (no secrets inside - secrets come from env vars named in it).

Autotest root (WORK) = the folder that holds qa.config.json; it is found from env QA_HOME, else the current directory, else its parents.
Every URL / path / flag / project rule used by the framework comes from here, so the same code tests any web project.

Environment overrides (all optional):
  QA_HOME            autotest root
  QA_CONFIG          another config file (e.g. qa.config.staging.json) instead of qa.config.json
  QA_OUT             output folder of the current run (set by tools.runs / autorun); default test_execution_output/adhoc/<date>
  QA_BASE_URL        URL of the app (another slot / environment)
  QA_ALLOW_EXTERNAL  1 = actions of risk EXTERNAL are allowed (only after the user's OK at the confirmation gate)
  QA_SUITE_DIR       catalog folder (default test_suite; sandbox: test_suite_<name>)
  QA_RESULTS_DIR     read/write results from another folder (tools.runs consolidate)"""
import datetime, json, os
from pathlib import Path

KIT = Path(__file__).resolve().parents[1]            # .../qakit


def _find_work():
    if os.environ.get("QA_HOME"):
        return Path(os.environ["QA_HOME"]).resolve()
    here = Path.cwd().resolve()
    for d in [here, *here.parents]:
        if (d / "qa.config.json").exists():
            return d
    return KIT.parent                                 # qakit lives inside the autotest root


WORK = _find_work()

DEFAULTS = {
    "project": "project",
    "language": {"talk": "vi", "excel": ["vi"], "testcase": "vi"},
    "app": {"url": "http://localhost:3000", "test_account": ""},
    "browser": {"mode": "launch", "channel": "chromium", "cdp_port": 9222, "profile_dir": "browser_profile", "headless": False,
                "viewport": [1600, 900], "slow_mo": 0},
    "auth": {"type": "none"},
    "matrix": {"envs": {}, "roles": {}, "default_env": "default", "default_role": "default"},
    "roles": {"switch": "none"},
    "envs": {"switch": "none"},
    "data": {"test_prefix": "ZZ_AUTOTEST_", "snapshot": "none", "hooks": "project.hooks"},
    "testcases": {"sources": ["testcases/*.xlsx"], "columns": {}, "version_labels": ["Version", "Phiên bản", "バージョン", "版数"]},
    "output": {"dir": "test_execution_output"},
    "source": {"use": True, "repos": {}},
}


def _merge(a, b):
    out = dict(a)
    for k, v in (b or {}).items():
        out[k] = _merge(a[k], v) if isinstance(v, dict) and isinstance(a.get(k), dict) else v
    return out


class Cfg:
    def __init__(self):
        p = Path(os.environ.get("QA_CONFIG") or (WORK / "qa.config.json"))
        if not p.is_absolute():
            p = WORK / p
        raw = json.load(open(p, encoding="utf-8")) if p.exists() else {}
        self.path = p
        self.raw = _merge(DEFAULTS, raw)
        r = self.raw
        self.name = r["project"]
        self.base_url = os.environ.get("QA_BASE_URL", r["app"]["url"]).rstrip("/")
        self.account = r["app"].get("test_account", "")
        b = r["browser"]
        self.browser = b
        self.cdp_port = int(os.environ.get("QA_CDP_PORT", b.get("cdp_port", 9222)))
        self.show_filename = os.environ.get("QA_SHOW_FILENAME", "0") == "1"
        self.prefix = r["data"]["test_prefix"]
        self.langs = list(r["language"]["excel"])
        self.talk = r["language"]["talk"]
        self.use_source = bool(r["source"].get("use", True))   # False = black-box: agent / tools never read the app source
        out = os.environ.get("QA_OUT") or str(WORK / r["output"]["dir"] / "adhoc" / datetime.date.today().strftime("%Y%m%d"))
        self.out = Path(out) if Path(out).is_absolute() else WORK / out        # a relative QA_OUT is relative to the autotest folder: every path written down is complete
        if os.environ.get("QA_OUT"):
            self.ensure()                      # ad-hoc folder: created on first write only (read-only tools leave no empty folders)

    def ensure(self):
        for d in (self.excel_dir, self.evidence_root, self.reports_dir, self.data_dir, self.safety_dir, self.logs_dir):
            d.mkdir(parents=True, exist_ok=True)

    def url(self, path=""):
        return self.base_url + path

    def secret(self, env_name):
        """Value of an env var named in the config (passwords, tokens). Never printed by the framework."""
        v = os.environ.get(env_name or "", "")
        if not v:
            raise RuntimeError(f"missing environment variable {env_name} (named in qa.config.json)")
        return v

    suite = None   # set by Suite(name): evidence/base/<suite>/<ID> and data/results_<suite>.json

    def set_suite(self, name):
        self.suite = name

    @property
    def output_root(self): return WORK / self.raw["output"]["dir"]
    @property
    def evidence_root(self): return self.out / "evidence" / "base"                  # annotated originals
    @property
    def evidence_dir(self): return self.evidence_root / self.suite if self.suite else self.evidence_root
    def evidence_lang_dir(self, lang): return self.out / "evidence" / lang             # re-captioned copies
    @property
    def excel_dir(self): return self.out / "excel"                                   # deliverables
    @property
    def reports_dir(self): return self.out / "reports"                               # for people (.md)
    @property
    def data_dir(self): return self.out / "data"                                     # for tools: results_*.json, preflight.json
    @property
    def safety_dir(self): return self.out / "safety"                                 # registry, snapshots, role backups
    @property
    def snapshots_dir(self): return self.safety_dir / "snapshots"
    @property
    def logs_dir(self): return self.out / "logs"
    @property
    def tmp_dir(self): return self.out / "tmp"


cfg = Cfg()
