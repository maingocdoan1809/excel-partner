# -*- coding: utf-8 -*-
"""COMMAND CHECK for /system-test (read-only). Validates the user's command against the fixed grammar and prints the normalised form.
  python -m qakit.tools.cmd <command> [sub] [key=value ...]          (tokens as the shell split them; one quoted string is split too)
  python -m qakit.tools.cmd --help [<command>]                      (usage of every / one command)
Exit 0 = valid: prints the canonical command and, for scope commands, the `tools.scope` tokens (also as JSON with --json).
Exit 2 = invalid: prints every problem with the correct form, then the usage of that command. Nothing is guessed:
bare values, unknown keys, repeated keys, values outside a fixed list, and missing required keys are all rejected."""
import json, re, shlex, sys
from pathlib import Path
from ..core.config import WORK, cfg
from . import catalog as cat_mod

sys.stdout.reconfigure(encoding="utf-8")

_M = cfg.raw.get("matrix", {})
ENVS = tuple(k for k in _M.get("envs", {}) if not k.startswith("_"))       # keys of qa.config.json matrix.envs / matrix.roles
ROLES = tuple(k for k in _M.get("roles", {}) if not k.startswith("_"))
SELECT = ("all", "changed", "failed")
STATUSES = ("PASS", "FAIL", "PENDING", "BLOCKED", "NOT_EXECUTED", "NOT EXECUTED")
VIEWS = ("summary", "full", "excel")
YESNO = ("yes", "no")
RUNS = cfg.output_root / "runs"

# key -> (meaning, value hint). Order = order shown in usage.
KEYS = {
    "file": ("file testcase: tên file .xlsx đầy đủ hoặc đường dẫn (không viết tắt)", "<tên file>.xlsx | testcases/<tên file>.xlsx"),
    "id": ("ID testcase", "ORD-001,ORD-002"),
    "suite": ("suite automation", "<tên suite>"),
    "select": ("tập case có sẵn", "|".join(SELECT)),
    "status": ("case có trạng thái đó trong một lần chạy", "FAIL,PENDING"),
    "run": ("run_id (thư mục runs/ trong output; bỏ trống = mới nhất)", "YYYYMMDD_HHMMSS_<env>_<nhãn>"),
    "tag": ("LỌC: phân loại / loại chứa chữ đó, hoặc tiền tố ID", "\"Phân quyền\" | ORD"),
    "env": ("LỌC slot môi trường (khóa matrix.envs)", "|".join(ENVS) or "(dự án không khai báo)"),
    "role": ("LỌC slot vai trò (khóa matrix.roles)", "|".join(ROLES) or "(dự án không khai báo)"),
    "exclude": ("LOẠI TRỪ: ID, tên file đầy đủ, suite, tag, env, role", "ORD-010,ORD-011"),
    "sandbox": ("catalog riêng test_suite_<tên> (demo / thử nghiệm)", "demo"),
    "url": ("URL môi trường cần test (mặc định app.base_url của config)", "https://..."),
    "view": ("cách xem kết quả", "|".join(VIEWS)),
    "base": ("mốc so sánh: run_id (result) hoặc build (suite diff)", "<run_id> | YYYYMMDD_HHMMSS"),
    "note": ("ghi chú cho lần nạp testcase", "\"spec v1.1\""),
    "ref": ("thay đổi source: task / PR / branch / commit / range", "PR-1234 | feature/x | abc123..def456"),
    "execute": ("chạy các case bị ảnh hưởng sau khi phân tích", "|".join(YESNO)),
    "desc": ("mô tả thay đổi bằng lời", "\"đổi nút Lưu ở màn đơn hàng\""),
}
SCOPE = ("file", "id", "suite", "select", "status", "run", "tag", "env", "role", "exclude", "sandbox")
SELECTORS = ("file", "id", "suite", "select", "status")

# command -> (description, allowed keys, required keys, needs a selector, sub-commands {sub: (desc, keys, required, needs selector)})
CMDS = {
    "": ("Tình trạng: catalog, lần chạy gần nhất, gợi ý bước tiếp", (), (), False),
    "check": ("Kiểm tra môi trường (Edge, phiên, DB thật, vai trò)", ("url",), (), False),
    "plan": ("Xem trước kế hoạch, không chạy", SCOPE, (), True),
    "run": ("Chạy thật (một phiên xác nhận trước khi ghi)", SCOPE + ("url",), (), True),
    "retry": ("Chạy lại FAIL + NOT EXECUTED của một lần chạy", ("run", "url", "sandbox"), (), False),
    "result": ("Kết quả + so với lần trước", ("run", "view", "base"), (), False),
    "update": ("Nạp file testcase mới / đã sửa vào Test Suite", ("note", "sandbox"), (), False),
    "clean": ("Dọn dữ liệu test (tiền tố data.test_prefix)", ("url",), (), False),
    "automate": ("Viết / sửa code test", SCOPE, (), True),
    "source": ("Source FE/BE đổi: tìm case ảnh hưởng, sửa test, chạy lại", ("ref", "url", "execute"), ("ref",), False),
}
SUITE_SUBS = {
    "show": ("Xem catalog (toàn bộ, theo file, hoặc từng ID)", ("file", "id", "suite", "sandbox"), (), False),
    "diff": ("Chênh lệch giữa hai lần nạp testcase", ("base", "sandbox"), (), False),
    "validate": ("Kiểm tra lỗi catalog / suite", ("sandbox",), (), False),
    "coverage": ("Độ phủ automation", ("sandbox",), (), False),
    "impact": ("Case bị ảnh hưởng bởi một thay đổi (chỉ phân tích)", ("ref", "desc"), (), False),
}
OLD = {"env-check": "check", "results": "result", "compare": "result", "report": "result view=full", "cleanup": "clean",
       "update-from-source": "source ref=<...>"}
EXAMPLES = {
    "": ["/system-test"],
    "check": ["/system-test check"],
    "plan": ["/system-test plan file=<tên file>.xlsx", "/system-test plan file=<tên file>.xlsx exclude=ORD-010", "/system-test plan select=changed"],
    "run": ["/system-test run file=<tên file>.xlsx", "/system-test run id=ORD-001,ORD-002", "/system-test run file=testcases/<file>.xlsx",
            "/system-test run select=failed run=<run_id>", "/system-test run file=<tên file>.xlsx sandbox=demo"],
    "retry": ["/system-test retry", "/system-test retry run=<run_id>"],
    "result": ["/system-test result", "/system-test result view=full", "/system-test result run=<run_id> base=<run_id>", "/system-test result view=excel"],
    "update": ["/system-test update note=\"spec v1.1\""],
    "clean": ["/system-test clean"],
    "automate": ["/system-test automate file=<tên file>.xlsx", "/system-test automate id=ORD-001,ORD-002"],
    "source": ["/system-test source ref=feature/x", "/system-test source ref=abc123..def456 execute=no"],
    "suite": ["/system-test suite show file=<tên file>.xlsx", "/system-test suite diff base=<build>", "/system-test suite coverage",
              "/system-test suite impact desc=\"đổi nút Lưu ở màn đơn hàng\""],
}


def usage(name=None):
    names = [name] if name is not None else list(CMDS) + ["suite"]
    out = ["Cú pháp: /system-test <lệnh> [tham số=giá trị ...]   (nhiều giá trị: dấu phẩy, không dấu cách; có dấu cách: đặt trong ngoặc kép)"]
    for n in names:
        rows = [(f"suite {s}", *v) for s, v in SUITE_SUBS.items()] if n == "suite" else [(n or "(trống)", *CMDS[n])]
        for label, desc, keys, req, sel in rows:
            ks = " ".join(f"{k}=" + ("" if k in req else "?") for k in keys)
            out.append(f"\n  {label:<16} {desc}" + (f"\n  {'':<16} tham số: {ks}" if keys else ""))
            if sel:
                out.append(f"  {'':<16} cần ít nhất một trong: {', '.join(k + '=' for k in SELECTORS)}")
            if req:
                out.append(f"  {'':<16} bắt buộc: {', '.join(k + '=' for k in req)}")
        for e in EXAMPLES.get(n, []):
            out.append(f"  {'':<16} ví dụ: {e}")
    if name is None or name in ("plan", "run", "automate", "suite"):
        out.append("\nTham số:")
        out += [f"  {k + '=':<9} {m}  ({h})" for k, (m, h) in KEYS.items()]
    return "\n".join(out)


def hint_bare(t):
    """A bare token -> the named form the user probably meant (shown as a suggestion, never applied)."""
    if t.startswith("-") and len(t) > 1:
        return f"exclude={t.lstrip('-')}"
    v = t.lower()
    if v in SELECT or v.split("@")[0] in SELECT + ("retry",):
        base, _, r = v.partition("@")
        return f"select={'failed' if base == 'retry' else base}" + (f" run={r}" if r else "")
    if v in ENVS:
        return f"env={v}"
    if v in cat_mod.sandboxes():
        return f"sandbox={v}"
    if v in ROLES:
        return f"role={v}"
    if v in VIEWS:
        return f"view={v}"
    if v == "no-run":
        return "execute=no"
    if re.fullmatch(r"\d{8}_\d{6}_.+|legacy(:.+)?", t):
        return f"run={t}"
    if t.lower().endswith(".xlsx") or "/" in t or "\\" in t:
        return f"file={t}"
    if all(re.fullmatch(r"\d{1,2}", x) for x in t.split(",")):
        return "file=<tên file .xlsx đầy đủ> (không nhận số file)"
    if all(re.fullmatch(r"[A-Za-z]+-\d+", x) for x in t.split(",")):
        return f"id={t}"
    return None


def check_value(k, v, errs):
    vs = [x.strip() for x in v.split(",")] if k not in ("note", "desc", "tag", "url", "ref") else [v]
    if any(x == "" for x in vs):
        errs.append(f"{k}=: giá trị trống hoặc thừa dấu phẩy ('{v}')")
        return v
    one = {"select": SELECT, "view": VIEWS, "execute": YESNO}
    many = {"env": ENVS, "role": ROLES}
    if k in many and not many[k]:
        errs.append(f"{k}=: qa.config.json không khai báo matrix.{k}s, tham số này không dùng được")
        return v
    if k in one:
        if len(vs) > 1 or vs[0] not in one[k]:
            errs.append(f"{k}={v}: chỉ nhận một trong {'|'.join(one[k])}")
    elif k in many:
        bad = [x for x in vs if x not in many[k]]
        if bad:
            errs.append(f"{k}={v}: không nhận {', '.join(bad)}; chỉ nhận {'|'.join(many[k])}")
    elif k == "status":
        bad = [x for x in vs if x.upper().replace("_", " ") not in STATUSES]
        if bad:
            errs.append(f"status={v}: không nhận {', '.join(bad)}; chỉ nhận PASS|FAIL|PENDING|BLOCKED|NOT_EXECUTED")
        vs = [x.upper().replace("_", " ") for x in vs]
    elif k in ("run", "base") and v != "latest":
        if k == "run" or re.fullmatch(r"\d{8}_\d{6}_.+", v):
            if not re.fullmatch(r"\d{8}_\d{6}_.+|legacy(:.+)?", v):
                errs.append(f"{k}={v}: run_id phải đầy đủ dạng YYYYMMDD_HHMMSS_<env>_<nhãn> (xem: python -m qakit.tools.runs list)")
            elif not v.startswith("legacy") and not (RUNS / v).is_dir():
                errs.append(f"{k}={v}: không có thư mục {RUNS / v}")
        elif not re.fullmatch(r"\d{8}_\d{6}", v):
            errs.append(f"base={v}: phải là run_id đầy đủ (result) hoặc tên build YYYYMMDD_HHMMSS (suite diff)")
    elif k == "sandbox" and v not in cat_mod.sandboxes():
        errs.append(f"sandbox={v}: không có test_suite_{v}/ (có: {', '.join(cat_mod.sandboxes()) or 'không có'})")
    elif k == "url" and not re.fullmatch(r"https?://\S+", v):
        errs.append(f"url={v}: phải bắt đầu bằng http:// hoặc https://")
    return ",".join(vs) if k not in ("note", "desc", "tag", "url", "ref") else v


def check_files(key, v, errs, sandbox=None):
    """Each value must be a path to a testcase .xlsx or its full file name (as listed by the catalog sources). Returns the file names
    (tools.scope matches by name). Shorthand (number, part of the name) is rejected with the full names it would have matched."""
    if sandbox:
        cat_mod.use(f"test_suite_{sandbox}")
    srcs = cat_mod.sources()                                  # relative paths matched by sources.json (new files included)
    names = {Path(s).name.lower(): Path(s).name for s in srcs}
    out = []
    for x in [y.strip() for y in v.split(",")]:
        if "/" in x or "\\" in x:
            p = Path(x)
            p = p if p.is_absolute() else WORK / p
            if not x.lower().endswith(".xlsx"):
                errs.append(f"{key}={x}: đường dẫn phải trỏ tới file .xlsx")
            elif not p.is_file():
                errs.append(f"{key}={x}: không tìm thấy file (đường dẫn tính từ {WORK})")
            elif cat_mod.rel(p) not in srcs:
                errs.append(f"{key}={x}: file không thuộc nguồn testcase ({cat_mod.SOURCES.name}: {', '.join(json.load(open(cat_mod.SOURCES, encoding='utf-8'))['sources']) if cat_mod.SOURCES.exists() else 'mặc định'})")
            out.append(p.name)
        elif x.lower().endswith(".xlsx"):
            if x.lower() not in names:
                near = [n for n in names.values() if x.lower()[:-5] in n.lower()]
                errs.append(f"{key}={x}: không có file tên này trong nguồn testcase" + (f" (gần giống: {', '.join(near)})" if near else ""))
            out.append(names.get(x.lower(), x))
        elif "." in x.rsplit("/", 1)[-1] and not x.lower().endswith(".xlsx"):
            errs.append(f"{key}={x}: chỉ nhận file testcase .xlsx")
            out.append(x)
        else:
            cands = [n for n in names.values() if (x.isdigit() and f"_{int(x):02d}_" in n) or (not x.isdigit() and x.lower() in n.lower())]
            errs.append(f"{key}={x}: không nhận viết tắt, dùng tên file đầy đủ (.xlsx) hoặc đường dẫn" +
                        (f"; ví dụ: {key}={cands[0]}" if len(cands) == 1 else f"; khớp: {', '.join(cands)}" if cands else ""))
            out.append(x)
    return ",".join(out)


def check_exclude(v, errs, sandbox=None):
    """exclude= mixes IDs, file names, suites, tags, env / role. File values follow check_files; a number or a part of a file name that is
    not an ID / suite / env / role is shorthand -> rejected."""
    out = []
    if sandbox:
        cat_mod.use(f"test_suite_{sandbox}")
    names = [Path(s).name for s in cat_mod.sources()]
    cat = cat_mod.load() or {}
    suites = set(cat.get("suites", {}))
    for x in [y.strip() for y in v.split(",")]:
        if x.lower().endswith(".xlsx") or "/" in x or "\\" in x:
            out.append(check_files("exclude", x, errs, sandbox))
        elif x.isdigit() or (x not in ENVS and x not in ROLES and x not in suites and not re.fullmatch(r"[A-Za-z]+-\d+", x)
                             and any(x.lower() in n.lower() for n in names)):
            out.append(check_files("exclude", x, errs, sandbox))         # reports the shorthand + full names
        else:
            out.append(x)
    return ",".join(out)


def parse(tokens):
    errs, toks = [], list(tokens)
    if len(toks) == 1 and " " in toks[0].strip():
        toks = shlex.split(toks[0])
    name = toks.pop(0).lower() if toks and "=" not in toks[0] else ""
    sub = None
    if name in OLD:
        return {"ok": False, "command": name, "errors": [f"Lệnh '{name}' đã bỏ, dùng: /system-test {OLD[name]}"]}
    if name and name not in CMDS and name != "suite":
        h = hint_bare(name)
        return {"ok": False, "command": "", "errors": [f"Không có lệnh '{name}'" + (f" (có phải: run {h}?)" if h else "") +
                                                        f". Lệnh: {', '.join(n for n in CMDS if n)}, suite"]}
    if name == "suite":
        sub = toks.pop(0).lower() if toks and "=" not in toks[0] else None
        if sub not in SUITE_SUBS:
            return {"ok": False, "command": "suite", "errors": [f"suite cần một trong: {', '.join(SUITE_SUBS)}" +
                                                                (f" (nhận được '{sub}')" if sub else "")]}
        desc, allowed, req, needs = SUITE_SUBS[sub]
    else:
        desc, allowed, req, needs = CMDS[name]
    params, raw = {}, {}                                   # raw = as typed (canonical line), params = normalised (scope tokens)
    for t in toks:
        if "=" not in t:
            h = hint_bare(t)
            errs.append(f"'{t}': thiếu tên tham số" + (f", viết: {h}" if h else f"; tham số của lệnh này: {', '.join(k + '=' for k in allowed) or '(không có)'}"))
            continue
        k, v = t.split("=", 1)
        k = k.lstrip("-").lower()
        if k not in KEYS:
            errs.append(f"'{k}=' không phải tham số" + (f"; tham số của lệnh này: {', '.join(x + '=' for x in allowed)}" if allowed else "; lệnh này không có tham số"))
            continue
        if k not in allowed:
            errs.append(f"'{k}=' không dùng với lệnh '{(name + ' ' + sub) if sub else name or '(trống)'}'" +
                        (f"; tham số hợp lệ: {', '.join(x + '=' for x in allowed)}" if allowed else "; lệnh này không có tham số"))
            continue
        if k in params:
            errs.append(f"'{k}=' lặp lại; gộp các giá trị bằng dấu phẩy: {k}={params[k]},{v}")
            continue
        raw[k] = v.strip()
        params[k] = check_value(k, raw[k], errs)
    sb = params.get("sandbox") if "sandbox" in params and not any("sandbox=" in e for e in errs) else None
    if "file" in params:
        params["file"] = check_files("file", raw["file"], errs, sb)
    if "exclude" in params:
        params["exclude"] = check_exclude(raw["exclude"], errs, sb)
    for k in req:
        if k not in params:
            errs.append(f"thiếu tham số bắt buộc {k}=")
    if needs and not any(k in params for k in SELECTORS):
        errs.append(f"cần chọn phạm vi bằng ít nhất một trong {', '.join(k + '=' for k in SELECTORS)} (không mặc định chạy tất cả; muốn tất cả: select=all)")
    if "run" in params and name in ("plan", "run", "automate") and params.get("select") != "failed" and "status" not in params:
        errs.append("run=<run_id> trong phạm vi chỉ đi cùng select=failed hoặc status=")
    if sub == "impact" and not ({"ref", "desc"} & set(params)):
        errs.append("suite impact cần ref= hoặc desc=")
    label = " ".join(x for x in ("/system-test", name, sub) if x)
    canon = " ".join([label] + [f'{k}="{v}"' if " " in v else f"{k}={v}" for k, v in raw.items()])
    res = {"ok": not errs, "command": name, "sub": sub, "params": params, "canonical": canon, "errors": errs}
    if not errs and needs:
        res["scope_tokens"] = scope_tokens(params)
    if not errs and name == "retry":
        res["scope_tokens"] = ["failed" + (f"@{params['run']}" if "run" in params else "")] + ([f"sandbox={params['sandbox']}"] if "sandbox" in params else [])
    return res


def scope_tokens(p):
    """Named parameters -> the long-form tokens tools.scope accepts (no bare values, so nothing is classified by guess)."""
    at = f"@{p['run']}" if "run" in p else ""
    t = [f"{k}={p[k]}" for k in ("file", "id", "suite") if k in p]
    if "select" in p:
        t.append(p["select"] + (at if p["select"] == "failed" else ""))
    if "status" in p:
        t.append(f"status={p['status'].replace(' ', '_')}{at}")
    t += [f"{k}={p[k]}" for k in ("tag", "env", "role", "exclude", "sandbox") if k in p]
    return t


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    as_json = "--json" in argv
    argv = [a for a in argv if a != "--json"]
    if argv[:1] in (["--help"], ["-h"]):
        n = argv[1].lower() if len(argv) > 1 else None
        print(usage(n if n in CMDS or n == "suite" else None))
        return 0
    r = parse(argv)
    if as_json:
        print(json.dumps(r, ensure_ascii=False, indent=1))
    elif r["ok"]:
        print(f"OK  {r['canonical']}")
        if r.get("scope_tokens"):
            print("scope: " + " ".join(shlex.quote(x) for x in r["scope_tokens"]))
    else:
        print("Lệnh không hợp lệ, chưa làm gì:")
        print("\n".join(f"  - {e}" for e in r["errors"]))
        cmdname = r.get("command")
        print("\n" + usage(cmdname if cmdname in CMDS or cmdname == "suite" else None))
    return 0 if r["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
