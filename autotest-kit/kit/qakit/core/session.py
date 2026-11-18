# -*- coding: utf-8 -*-
"""Login, role and environment switching through the adapters chosen in qa.config.json. Nothing here is project specific.

auth.type
  none     the app needs no login
  form     username / password form: {"login_url", "username_env", "password_env", "user_selector", "password_selector",
           "submit_selector", "logout_url" | "logout_selector", "logged_in": {"selector": css} | {"url_not_contains": "/login"}}
  manual   a human logs in once in the test browser (SSO / MFA); the framework only checks `logged_in` and reports PENDING when not
roles.switch (how the runner becomes another role of the matrix)
  none     single account: cases of other roles become PENDING ("role switching not configured")
  accounts {"accounts": {"<role>": {"username_env": "...", "password_env": "..."}}}: log out, log in with that role's account (auth form)
  command  {"command": "python tools/set_role.py --role {role}", "relogin": true, "settle_seconds": 0}: a project script changes the role
           of the test account, then the framework logs out / in so the app reloads the permissions
envs.switch (how the runner switches a matrix environment)
  none     only the default environment
  command  {"command": "python tools/set_env.py --env {env}"} (restored with the default env at the end)
Labels: matrix.envs / matrix.roles map a key to the words that identify it in the testcase role column; the first word is the label
printed in evidence captions, e.g. "production": ["本番", "production"] -> [本番·admin]."""
import shlex, subprocess, time
from .config import cfg, WORK

A = cfg.raw["auth"]
M = cfg.raw["matrix"]
ENV_LABEL = {k: (v[0] if v else k) for k, v in M.get("envs", {}).items()}
ENV_LABEL.setdefault(M.get("default_env", "default"), M.get("default_env", "default"))
_state = {"role": M.get("default_role", "default"), "env": M.get("default_env", "default")}


class NeedsHuman(RuntimeError):
    """Login / switch that only a human can do (MFA, SSO approval, missing adapter). The runner records PENDING."""


# ---------------------------------------------------------------------------------------------------------------------------- login
def logged_in(pg):
    li = A.get("logged_in") or {}
    if A.get("type", "none") == "none":
        return True
    if li.get("selector"):
        try:
            return pg.locator(li["selector"]).first.is_visible()
        except Exception:
            return False
    if li.get("url_not_contains"):
        return li["url_not_contains"] not in pg.url
    return True


def login(pg, account=None):
    """Log in with the default account (auth.username_env / password_env) or `account` = {"username_env", "password_env"}."""
    t = A.get("type", "none")
    if t == "none":
        return
    if t == "manual":
        raise NeedsHuman("login required: log in once in the test browser (SSO / MFA), then rerun")
    if t == "session":
        raise NeedsHuman(f"the saved sign-in '{A.get('session_name', '')}' has expired: sign in again at Settings > Testing > Sign-ins ('Sign in again'), then rerun. "
                         "Do NOT switch the browser to cdp or type a password.")
    if t != "form":
        raise NeedsHuman(f"unknown auth.type {t}")
    acc = account or A
    pg.goto(cfg.url(A.get("login_url", "/login")))
    pg.wait_for_load_state("domcontentloaded")
    pg.locator(A["user_selector"]).first.fill(cfg.secret(acc["username_env"]))
    pg.locator(A["password_selector"]).first.fill(cfg.secret(acc["password_env"]))
    pg.locator(A["submit_selector"]).first.click()
    pg.wait_for_load_state("networkidle")
    pg.wait_for_timeout(int(A.get("settle_ms", 1000)))
    if not logged_in(pg):
        raise RuntimeError("login failed (check credentials env vars and auth selectors in qa.config.json)")


def logout(pg):
    if A.get("logout_url"):
        pg.goto(cfg.url(A["logout_url"]))
    elif A.get("logout_selector"):
        pg.locator(A["logout_selector"]).first.click()
    else:
        pg.context.clear_cookies()
    pg.wait_for_timeout(800)


def ensure_session(pg):
    """True when a login had to be performed. Raises NeedsHuman when only a human can log in."""
    if logged_in(pg):
        return False
    login(pg, _account(_state["role"]))
    return True


# ---------------------------------------------------------------------------------------------------------------------------- roles
R = cfg.raw["roles"]


def _account(role):
    acc = R.get("accounts", {}).get(role)
    return acc if R.get("switch") == "accounts" and acc else None


def _run(cmd_tpl, **kw):
    cmd = cmd_tpl.format(**kw)
    p = subprocess.run(shlex.split(cmd, posix=False), cwd=WORK, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode:
        raise RuntimeError(f"command failed ({p.returncode}): {cmd}\n{(p.stdout + p.stderr)[-400:]}")
    return p.stdout


def switch_role(pg, role):
    """Become `role` (a key of matrix.roles, or "home" = default role)."""
    role = M.get("default_role", "default") if role == "home" else role
    if role == _state["role"]:
        return
    mode = R.get("switch", "none")
    if mode == "none":
        raise NeedsHuman(f"role switching not configured (roles.switch = none): cannot run as '{role}'")
    if mode == "accounts":
        acc = _account(role) or (A if role == M.get("default_role") else None)
        if not acc:
            raise NeedsHuman(f"no test account for role '{role}' in roles.accounts")
        logout(pg)
        login(pg, acc)
    elif mode == "command":
        _run(R["command"], role=role)
        time.sleep(float(R.get("settle_seconds", 0)))
        if R.get("relogin", True):
            logout(pg)
            ensure_session(pg)
    _state["role"] = role
    print("role ->", role)


def role_command(role):
    """Command a human can run to restore a role (printed when the runner cannot restore it)."""
    if R.get("switch") == "command":
        return R["command"].format(role=M.get("default_role") if role == "home" else role)
    return f"log in again with the default test account ({cfg.account})"


# ---------------------------------------------------------------------------------------------------------------------------- envs
E = cfg.raw["envs"]


def switch_env(env, roles=None):
    if env == _state["env"]:
        return
    if E.get("switch", "none") == "none":
        raise NeedsHuman(f"environment switching not configured (envs.switch = none): cannot run in '{env}'")
    _run(E["command"], env=env)
    _state["env"] = env
    print("env ->", env)


def role_env_state():
    return dict(_state)


def restore_env(orig):
    if orig.get("env") and orig["env"] != _state["env"]:
        switch_env(orig["env"])
    return dict(_state)


def backup_roles():
    """Projects whose role switching edits shared data can back it up in project hooks (hooks.backup_roles()); default: nothing."""
    from . import hooks
    fn = getattr(hooks.module(), "backup_roles", None)
    return fn(cfg.safety_dir) if fn else None
