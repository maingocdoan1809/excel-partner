# -*- coding: utf-8 -*-
"""SCOPE RESOLVER -> RUN PLAN (read-only; needs a catalog from tools.catalog build). Printing the plan IS the dry-run.
  python -m qakit.tools.scope <scope...> [--out plan.json] [--brief]   (--brief = reviewer view for the confirmation gate)
Short syntax (what /system-test users type; comma = several values, no spaces):
  09 | 09,02                 IDs of those workbooks (number = the _NN_ part of the file name)
  quan_ly_user               IDs of the workbooks whose name contains the text
  USR-001,HOME-026           explicit IDs (deprecated IDs are allowed only this way, with a warning)
  user_mgmt                  every case of a suite
  all                        every active catalog ID (default when no base token is given)
  changed                    IDs new/changed in the latest catalog build or still needing review
  failed[@run]               FAIL + NOT EXECUTED of a run (default latest; see tools.runs); retry = failed
  status=FAIL,PENDING[@run]  IDs with those statuses in a run
  tag=<text>,...             FILTER: 分類 / タイプ substring or ID prefix (USR, HOME-0)
  env=<env>  role=<role>     FILTER on (environment, role) slots of matrix suites (keys / keywords of qa.config.json matrix) -> --targets
  -<v>,...                   EXCLUDE by ID, file number/name, tag, suite, or env / role (key or keyword of matrix.envs / matrix.roles)
  demo                       sandbox: use the catalog test_suite_demo/ (any test_suite_<name>/ folder works the same)
A bare value matching no file / ID / suite is reported as unknown (exit 2), never guessed. Long forms file= id= suite= exclude= sandbox= and
--key=value are still accepted. Base tokens are combined as a union, filters narrow it, exclude removes. IDs without automation are listed
as NOT AUTOMATED (never run). Catalog freshness: a source workbook saved after the last `catalog build` (or a new / missing one) that
touches the scope -> the plan is marked STALE and the exit code is 2 (run `catalog build` first)."""
import json, re, sys
from pathlib import Path
from ..core.config import WORK, cfg
from . import catalog as cat_mod

sys.stdout.reconfigure(encoding="utf-8")
from ..core.config import cfg
_M = cfg.raw["matrix"]
ENV_ALIAS = {w: k for k, ws in _M.get("envs", {}).items() for w in [k, *ws]}       # every key / keyword of matrix.envs -> key
ROLE_ALIAS = {w: k for k, ws in _M.get("roles", {}).items() for w in [k, *ws]}
ROLES = tuple(_M.get("roles", {}))


def vals(tok):
    return [v.strip() for v in tok.split("=", 1)[1].split(",") if v.strip()] if "=" in tok else []


def file_match(f, v):
    name = Path(f).name
    return (v.isdigit() and f"_{int(v):02d}_" in name) or (not v.isdigit() and v.lower() in name.lower())


def tag_match(c, v):
    v2 = v.lower()
    return v2 in c.get("cls", "").lower() or v2 in c.get("typ", "").lower() or c["id"].lower().startswith(v2)


KEYWORDS = ("all", "changed", "failed", "retry", "retry-failed")


def normalize(tokens):
    """Short syntax -> canonical tokens. `-X,Y` = exclude=X,Y · bare word: sandbox name (test_suite_<name>) -> sandbox; otherwise each comma
    value is classified: number -> file, known ID -> id, suite name -> suite, else a file-name substring -> file (unknown when nothing matches).
    Returns (canonical tokens, sandbox name or None). key=value tokens and `--key=value` pass through unchanged."""
    out, sandbox = [], None
    for t in tokens:
        if t.startswith("-") and not t.startswith("--") and len(t) > 1:
            out.append("exclude=" + t[1:])
            continue
        t = t.lstrip("-")
        k = t.split("=", 1)[0].lower()
        if k == "sandbox" and "=" in t:
            sandbox = t.split("=", 1)[1]
            continue
        if "=" in t or k.split("@", 1)[0] in KEYWORDS or k in ("execute", "run", "plan", "dry-run"):
            out.append(t)
            continue
        if t.lower() in cat_mod.sandboxes():
            sandbox = t.lower()
            continue
        out.append("_bare=" + t)                          # classified after the catalog is loaded
    return out, sandbox


def classify(v, cat):
    if v.isdigit():
        return "file"
    if v in cat["cases"] or any(v in s["ids"] for s in cat["suites"].values()):
        return "id"
    if v in cat["suites"]:
        return "suite"
    if any(file_match(f, v) for f in cat["sources"]):
        return "file"
    return None


def resolve(tokens):
    typed = list(tokens)                                      # as the user wrote them (plan title, run index)
    tokens, sandbox = normalize(tokens)
    if sandbox:
        if not (WORK / f"test_suite_{sandbox}" / "sources.json").exists():
            sys.exit(f"sandbox {sandbox}: test_suite_{sandbox}/ does not exist")
        cat_mod.use(f"test_suite_{sandbox}")
    cat = cat_mod.load()
    if not cat:
        sys.exit("no catalog: run `python -m qakit.tools.catalog build` first")
    cases, suites = cat["cases"], cat["suites"]
    canon = []
    for t in tokens:                                          # bare values -> file= / id= / suite=
        if not t.startswith("_bare="):
            canon.append(t)
            continue
        vs = vals(t)
        kinds = {v: classify(v, cat) for v in vs}
        for kind in ("file", "id", "suite"):
            if any(k == kind for k in kinds.values()):
                canon.append(f"{kind}=" + ",".join(v for v in vs if kinds[v] == kind))
        canon += [f"?{v}" for v in vs if kinds[v] is None]
    tokens = canon
    toks = [t.lstrip("-") for t in tokens if t.lstrip("-").lower() not in ("execute", "run", "plan", "dry-run")]   # verbs of the skill command
    toks = ["failed" + t[t.index("@"):] if t.lower().split("@")[0] in ("retry", "retry-failed") and "@" in t else
            "failed" if t.lower() in ("retry", "retry-failed") else t for t in toks]
    base, warn, unknown = {}, [], []                           # id -> preferred file (for suite choice)
    suite_only = {}                                            # suite -> [ids] (spec-less / script suites)
    filters_tag, envs, roles, excl = [], set(), set(), []
    has_base = False
    for t in toks:
        k = t.split("=", 1)[0].split("@", 1)[0].lower()
        if k == "all":
            has_base = True
            base.update({i: c["file"] for i, c in cases.items() if c["status"] != "deprecated"})
        elif k == "file":
            has_base = True
            for v in vals(t):
                hit = {i: (c["file"] if file_match(c["file"], v) else next(f for f in c.get("also_in", []) if file_match(f, v)))
                       for i, c in cases.items() if c["status"] != "deprecated" and (file_match(c["file"], v) or any(file_match(f, v) for f in c.get("also_in", [])))}
                if not hit:
                    unknown.append(f"file={v}")
                base.update(hit)
        elif k == "id":
            has_base = True
            for v in vals(t):
                if v in cases:
                    base[v] = cases[v]["file"]
                    if cases[v]["status"] == "deprecated":
                        warn.append(f"{v} is deprecated (no longer in the source workbook)")
                elif any(v in s["ids"] for s in suites.values()):
                    for n, s in suites.items():
                        if v in s["ids"]:
                            suite_only.setdefault(n, []).append(v)
                            break
                else:
                    unknown.append(v)
        elif k == "suite" or (k in suites and "=" not in t):
            has_base = True
            for n in (vals(t) if k == "suite" else [k]):
                if n not in suites:
                    unknown.append(f"suite={n}")
                    continue
                if suites[n]["spec"] in cat["sources"]:
                    base.update({i: suites[n]["spec"] for i in suites[n]["ids"] if i in cases})
                else:
                    suite_only[n] = list(suites[n]["ids"]) or ["(whole suite)"]
        elif k == "changed":
            has_base = True
            base.update({i: c["file"] for i, c in cases.items() if c["status"] in ("new", "changed") or c.get("needs_review")})
        elif k in ("failed", "status"):
            has_base = True
            from . import runs
            run = t.split("@", 1)[1] if "@" in t else "latest"
            want = {"FAIL", "NOT EXECUTED"} if k == "failed" else {v.upper() for v in vals(t.split("@", 1)[0])}
            for (s, i), r in runs.results_of(run).items():
                if r.get("status") in want:
                    if i in cases:
                        base[i] = suites[s]["spec"] if s in suites and suites[s]["spec"] else cases[i]["file"]
                    else:
                        suite_only.setdefault(s, []).append(i)
        elif k == "tag":
            filters_tag += vals(t)
        elif k == "env":
            envs |= {ENV_ALIAS.get(v, v) for v in vals(t)}
        elif k == "role":
            roles |= {ROLE_ALIAS.get(v, v) for v in vals(t)}
        elif k == "exclude":
            excl += vals(t)
        else:
            unknown.append(t.lstrip("?"))
    if not has_base:
        base.update({i: c["file"] for i, c in cases.items() if c["status"] != "deprecated"})
    if filters_tag:
        base = {i: f for i, f in base.items() if any(tag_match(cases[i], v) for v in filters_tag)}
    excluded = []
    for v in excl:
        if ENV_ALIAS.get(v) or ROLE_ALIAS.get(v):
            continue
        drop = [i for i, f in base.items() if i == v or file_match(cases[i]["file"], v) or tag_match(cases[i], v)
                or any(a["suite"] == v for a in cases[i].get("automation", []))]
        for i in drop:
            base.pop(i)
            excluded.append(f"{i} (-{v})")
        if v in suite_only:
            excluded += [f"{i} (-{v})" for i in suite_only.pop(v)]
    ex_env = {ENV_ALIAS[v] for v in excl if v in ENV_ALIAS}
    ex_role = {ROLE_ALIAS[v] for v in excl if v in ROLE_ALIAS}

    def allowed(tg):
        e, r = tg.split(":")
        return (not envs or e in envs) and (not roles or r in roles) and e not in ex_env and r not in ex_role

    items, not_auto = {}, []
    for i, f in sorted(base.items()):
        c = cases[i]
        auto = c.get("automation", [])
        if not auto:
            not_auto.append(dict(id=i, file=Path(c["file"]).name, title=c["title"][:60]))
            continue
        pick = next((a for a in auto if suites.get(a["suite"], {}).get("spec") == f), None) or next((a for a in auto if a["primary"]), auto[0])
        s = pick["suite"]
        tg = [x for x in c["targets"] if allowed(x)]
        if suites[s]["mode"] == "matrix" and not tg:
            excluded.append(f"{i} (no (env, role) slot left after env/role filters)")
            continue
        it = items.setdefault(s, dict(suite=s, mode=suites[s]["mode"], spec=suites[s]["spec"], ids=[], targets=set()))
        it["ids"].append(i)
        it["targets"] |= set(tg)
    for s, ids in suite_only.items():
        it = items.setdefault(s, dict(suite=s, mode=suites[s]["mode"], spec=suites[s]["spec"], ids=[], targets=set()))
        it["ids"] += [i for i in ids if i not in it["ids"]]
    # catalog freshness: a source workbook saved after the build -> the plan would use old IDs / old content while the runner and the
    # Excel builder read the new file. Blocking when the stale file touches the plan (its cases, a suite spec, or a new file for all/changed).
    st = cat_mod.stale(cat)
    rel_files = {cases[i]["file"] for i in base} | set(base.values()) |                 {suites[s]["spec"] for s in list(items) + list(suite_only) if suites[s]["spec"]}   # primary file, file it was picked by, spec run
    wide = not has_base or any(t.split("=", 1)[0].split("@", 1)[0].lower() in ("all", "changed", "file") for t in toks)
    any_change = any(t.split("@", 1)[0].lower() == "changed" for t in toks)     # `changed` = diff vs the build -> every stale source matters
    stale_block = {f: w for f, w in st.items() if any_change or f in rel_files or (wide and w in ("new", "untracked"))
                   or any(t.lower().startswith("file=") and any(file_match(f, v) for v in vals(t)) for t in toks)}
    for f, w in st.items():
        if f not in stale_block:
            warn.append(f"catalog out of date for {f} ({cat_mod.STALE_TEXT[w]}), not used by this scope")
    filt = bool(envs or roles or ex_env or ex_role)
    out = []
    for it in items.values():
        it["targets"] = sorted(it["targets"]) if filt and it["mode"] == "matrix" else None
        if it["ids"] == ["(whole suite)"]:
            it["ids"] = []
        if (filt and it["mode"] != "matrix") and it["ids"]:
            warn.append(f"suite {it['suite']} is not a matrix suite: env/role filters do not apply (runs in the current role per case)")
        out.append(it)
    for s in out:
        if suites[s["suite"]]["todo"]:
            warn.append(f"suite {s['suite']} still has t.todo() steps -> those steps report NOT EXECUTED")
        if s["spec"] and not suites[s["suite"]]["spec_exists"]:
            warn.append(f"suite {s['suite']}: spec file missing ({s['spec']}) -> Excel cannot be built")
    for n, e in cat["suite_errors"].items():
        if n in items:
            warn.append(f"suite {n} cannot be loaded: {e}")
    return dict(scope=typed, resolved=tokens, sandbox=sandbox, catalog_dir=cat_mod.DIR.name, catalog_build=cat["build"], stale={f: cat_mod.STALE_TEXT[w] for f, w in stale_block.items()}, items=out, not_automated=not_auto, excluded=excluded, unknown=unknown, warnings=warn)


def command(it):
    a = ["python", "-m", f"suites.{it['suite']}.cases"] + (["--matrix"] if it["mode"] == "matrix" else []) + list(it["ids"])
    if it.get("targets"):
        a.append("--targets=" + ",".join(it["targets"]))
    return a


def render(plan):
    n = sum(len(i["ids"]) for i in plan["items"])
    L = [f"# Run plan (dry-run) — scope: {' '.join(plan['scope']) or 'all'}" + (f" · sandbox {plan['sandbox']}" if plan.get("sandbox") else ""), "",
         f"Catalog: {plan['catalog_build']} · will run **{n} case(s)** in {len(plan['items'])} suite(s) · NOT AUTOMATED: {len(plan['not_automated'])} · excluded: {len(plan['excluded'])}", "",
         "| Suite | Mode | Spec | Cases | Slots (env:role) | Command |", "|---|---|---|---|---|---|"]
    if plan.get("stale"):
        L[2:2] = ["## ⛔ Catalog out of date: this plan CANNOT run", "",
                  *[f"- `{Path(f).name}`: {w}" for f, w in plan["stale"].items()], "",
                  "→ Run `/system-test update" + (f" {plan['sandbox']}" if plan.get("sandbox") else "") + "` first, then `plan` / `run` again. "
                  "The plan below uses the old catalog, for reference only.", ""]
    for it in plan["items"]:
        ids = ", ".join(it["ids"][:12]) + (" …" if len(it["ids"]) > 12 else "") if it["ids"] else "(whole suite)"
        L.append(f"| {it['suite']} | {it['mode']} | {Path(it['spec']).name if it['spec'] else '—'} | {len(it['ids']) or 'all'}: {ids} | "
                 f"{', '.join(it['targets']) if it.get('targets') else 'from the role column' if it['mode'] == 'matrix' else 'current role / per case'} | "
                 f"`{' '.join(command(it)[:4])} …` |")
    if plan["not_automated"]:
        by = {}
        for x in plan["not_automated"]:
            by.setdefault(x["file"], []).append(x["id"])
        L += ["", "## NOT AUTOMATED (selected, no suite implements them: not run, never PASS)", *[f"- {f}: {len(v)} ({', '.join(v[:15])}{' …' if len(v) > 15 else ''})" for f, v in by.items()]]
    for k, t in (("excluded", "Excluded"), ("unknown", "Not recognised (check the scope)"), ("warnings", "Warnings")):
        if plan[k]:
            L += ["", f"## {t}", *[f"- {x}" for x in plan[k][:60]]]
    return "\n".join(L)


# ---- reviewer view (`--brief`): what the confirmation gate shows. Screens and actions in words; IDs / codes only as a reference column.
TXT = {
    "vi": dict(run="Sẽ chạy: {n} case trên {f} màn hình · môi trường {url}", slots="Vai trò / môi trường sẽ dùng", file="file",
               cols="| # | Kiểm tra gì | Chạy với | Dữ liệu test sẽ tạo | Mã |", cur="vai trò hiện tại", not_run="Không chạy",
               na="Chưa có code tự động, {s}: {k} case", ex="Loại trừ theo phạm vi bạn gõ: {k}", stale="⛔ Chưa chạy được: file testcase đã sửa sau lần nạp gần nhất",
               unk="⛔ Có giá trị không nhận ra trong phạm vi"),
    "en": dict(run="Will run: {n} case(s) on {f} screen(s) · environment {url}", slots="Roles / environments used", file="file",
               cols="| # | What is checked | Runs as | Test data created | ID |", cur="current role", not_run="Not run",
               na="No automation yet, {s}: {k} case(s)", ex="Excluded by your scope: {k}", stale="⛔ Cannot run: testcase file saved after the last update",
               unk="⛔ Values not recognised in the scope"),
}


def screen_name(f):
    """Display name of a testcase workbook for reviewers: test_suite/screens.json {"03": "Orders screen", "<part of name>": "..."} first,
    else from the file name (leading prefix up to _NN_ removed, underscores -> spaces)."""
    name = Path(f).stem
    sp = cat_mod.DIR / "screens.json"
    m = json.load(open(sp, encoding="utf-8")) if sp.exists() else {}
    for k, v in m.items():
        if k.startswith("_"):
            continue
        if (k.isdigit() and f"_{int(k):02d}_" in Path(f).name) or (not k.isdigit() and k.lower() in name.lower()):
            return v
    return re.sub(r"^.*?_\d{2}_", "", name).replace("_", " ").strip() or name


def slot_txt(tg):
    e, r = tg.split(":")
    el = (_M.get("envs", {}).get(e) or [e])[0]
    rl = (_M.get("roles", {}).get(r) or [r])[0]
    return f"{el} · {rl}" if _M.get("envs") else rl


def _notes(suite):
    """Optional suites/<suite>/notes.json {ID: {"data": ["what test data the case creates"]}} -> 'Test data created' column."""
    pth = WORK / "suites" / suite / "notes.json"
    return json.load(open(pth, encoding="utf-8")) if pth.exists() else {}


def brief(plan):
    """Markdown for a human reviewer: per screen (testcase file) the cases in words, the role / environment they run with, the test data
    they create; then what is NOT run and why. Source for the confirmation gate (skill references/confirm.md)."""
    T = TXT.get(cfg.talk, TXT["en"])
    cases = cat_mod.load()["cases"]
    by_file, slots_all, n = {}, set(), 0
    for it in plan["items"]:
        notes = _notes(it["suite"])
        for i in it["ids"]:
            c = cases.get(i, {})
            tg = [t for t in c.get("targets", []) if not it.get("targets") or t in it["targets"]] if it["mode"] == "matrix" else []
            slots_all |= set(tg)
            data = [d for d in notes.get(i, {}).get("data", []) if d and d != "—"]
            by_file.setdefault(c.get("file") or it.get("spec") or it["suite"], []).append((i, c, tg, data))
            n += 1
    L = ["## " + T["run"].format(n=n, f=len(by_file), url=cfg.base_url), ""]
    if slots_all and (_M.get("envs") or _M.get("roles")):
        L += [f"{T['slots']}: " + "; ".join(slot_txt(t) for t in sorted(slots_all)), ""]
    for f, rows in by_file.items():
        L += [f"### {screen_name(f)}  ·  {T['file']} `{Path(f).name}`", "", T["cols"], "|---|---|---|---|---|"]
        for k, (i, c, tg, data) in enumerate(rows, 1):
            what = (c.get("title") or "").replace("|", "/").replace("\n", " ")
            cls = " / ".join(x for x in (c.get("cls"), c.get("typ")) if x)
            L.append(f"| {k} | {what}" + (f" _({cls})_" if cls else "") + f" | {'; '.join(slot_txt(t) for t in tg) or T['cur']} | {'; '.join(data) or '—'} | {i} |")
        L.append("")
    if plan.get("not_automated") or plan.get("excluded"):
        L += ["## " + T["not_run"], ""]
        na = {}
        for x in plan.get("not_automated", []):
            na.setdefault(x["file"], []).append(f"{x['title'][:50]} ({x['id']})")
        for f, v in na.items():
            L.append("- " + T["na"].format(s=screen_name(f), k=len(v)) + " — " + "; ".join(v[:8]) + (" …" if len(v) > 8 else ""))
        ex = plan.get("excluded", [])
        if ex:
            L.append("- " + T["ex"].format(k=len(ex)) + " — " + "; ".join(
                f"{(cases.get(x.split(' ')[0], {}).get('title') or '')[:40]} ({x.split(' ')[0]})" for x in ex[:10]) + (" …" if len(ex) > 10 else ""))
        L.append("")
    if plan.get("stale"):
        L += ["## " + T["stale"], *[f"- {screen_name(f)} (`{Path(f).name}`)" for f in plan["stale"]], ""]
    if plan.get("unknown"):
        L += ["## " + T["unk"], *[f"- `{u}`" for u in plan["unknown"]], ""]
    return "\n".join(L)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    out = None
    if "--out" in argv:
        j = argv.index("--out")
        out = argv[j + 1]
        del argv[j:j + 2]
    use_brief = "--brief" in argv
    argv = [a for a in argv if a != "--brief"]
    plan = resolve(argv)
    print(brief(plan) if use_brief else render(plan))
    if out:
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        Path(out).write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")
        print("\nwritten", out)
    return 2 if plan["unknown"] or plan.get("stale") or not plan["items"] else 0


if __name__ == "__main__":
    sys.exit(main())
