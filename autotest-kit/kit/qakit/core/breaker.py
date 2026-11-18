# -*- coding: utf-8 -*-
"""Circuit breaker: stop a run when the SAME failure repeats too often (a broken login, a missing menu, a dead server ...).
Two failures are 'the same' when their signatures match: the first non-PASS step's text with digits / quoted values removed.
Limit = env QA_MAX_REPEAT (default 3). Only FAIL and NOT EXECUTED count; PENDING (needs a human) and BLOCKED do not."""
import os, re

COUNTED = ("FAIL", "NOT EXECUTED")


def limit():
    """env QA_MAX_REPEAT, else qa.config.json run.max_repeat, else 3."""
    try:
        from .config import cfg
        configured = cfg.raw.get("run", {}).get("max_repeat", 3)
    except Exception:
        configured = 3
    try:
        return max(2, int(os.environ.get("QA_MAX_REPEAT", configured)))
    except (TypeError, ValueError):
        return 3


def signature(steps, remark=""):
    st = next((s for s in steps if s.get("status") in COUNTED), None)
    text = (st.get("actual") or st.get("remark") or "") if st else remark
    text = re.sub(r"^script / runner error \(not a verdict on the application\):\s*", "", str(text))
    text = re.sub(r"[「『\"'].*?[」』\"']", "", text)
    text = re.sub(r"\d+", "#", text).lower()
    return re.sub(r"\s+", " ", text).strip()[:120]


class Breaker:
    def __init__(self):
        self.seen, self.tripped = {}, None

    def observe(self, status, steps, remark=""):
        """Call once per finished case. Returns the reason text when the run must stop, else None."""
        if status not in COUNTED:
            return None
        sig = signature(steps, remark)
        if not sig:
            return None
        self.seen[sig] = self.seen.get(sig, 0) + 1
        if self.seen[sig] >= limit():
            self.tripped = f"the same error repeated {self.seen[sig]} times: {sig}"
            return self.tripped
        return None
