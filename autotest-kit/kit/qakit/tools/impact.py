# -*- coding: utf-8 -*-
"""SOURCE-CHANGE IMPACT (read-only: git diff + catalog; nothing in the repos is written, no fetch).
  python -m qakit.tools.impact --range <repo>=<range> [--range <repo>=<range> ...] [--files <repo>:<path> ...] [--out <dir>]
<repo> = a key of test_suite/source_map.json "repos" (e.g. fe, be, app). <range> = A..B / A...B, or ONE commit / branch tip (also a merge
commit) = its first-parent diff.
Steps:
  1 changed files per repo (git diff --name-status)
  2 file -> testcase workbook through source_map.json rules: direct (every ID of the workbook) / cross (IDs whose 分類/タイプ match the
    rule tags) / broad (warning only); files matching no rule are listed as UNMAPPED (agent decides, then adds a rule)
  3 candidate IDs from the catalog, split into automated / NOT AUTOMATED
  4 UI strings removed by the diff (string literals, JSX text, locale values) that pages/ or suites/ still use -> automation likely broken
  5 docs/system_map sections that mention a changed file -> to refresh
Writes <out>/impact.json + impact.md (default: print only). The agent narrows the ID list by reading the diff.
source_map.json: {"repos": {"fe": "../frontend"}, "ignore": ["*.md"], "rules": [{"repo": "fe", "pattern": "src/orders/*", "files": ["03"],
"level": "direct", "area": "Orders", "map": "A §3"}]}   files = workbook numbers (_NN_) or parts of workbook names; "*" = every workbook."""
import argparse, fnmatch, json, re, subprocess, sys
from pathlib import Path
from ..core.config import WORK
from . import catalog as cat_mod

sys.stdout.reconfigure(encoding="utf-8")
MAP = WORK / "test_suite" / "source_map.json"
SYSMAP = WORK / "docs" / "system_map"
QUOTED = re.compile(r"""(["'`])((?:(?!\1).){2,80})\1""")
JSXTEXT = re.compile(r">([^<>{}]{2,80})<")
NONASCII = re.compile(r"[^\x00-\x7F][^\"'`<>{}\n]{0,78}")


def git(repo, *args):
    r = subprocess.run(["git", "-c", "core.quotepath=false", "-C", repo, *args], capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode:
        raise SystemExit(f"git {' '.join(args)} failed in {repo}: {r.stderr.strip()[:300]}  (refs missing? run `git fetch` in that repo)")
    return r.stdout


def norm_range(rng):
    if ".." in rng:
        return rng
    c = rng[:-2] if rng.endswith("^!") else rng
    return f"{c}^1..{c}"


def changed(repo, rng):
    out = []
    for line in git(repo, "diff", "--name-status", "-M", rng).splitlines():
        parts = line.split("\t")
        if len(parts) >= 2:
            out.append((parts[0][0], parts[-1].replace("\\", "/"), parts[1].replace("\\", "/") if parts[0].startswith("R") else None))
    return out


def _texts(line):
    out = set()
    for _, s in QUOTED.findall(line):
        if re.search(r"[A-Za-zÀ-￿]", s) and not re.match(r"^[\w./@-]+$", s) or re.search(r"[^\x00-\x7F]", s):
            out.add(s.strip())
    out |= {s.strip() for s in JSXTEXT.findall(line) if re.search(r"\w", s)}
    out |= {s.strip(" ・") for s in NONASCII.findall(line)}
    return {s for s in out if 2 <= len(s) <= 80}


def removed_strings(repo, rng, globs):
    """UI strings that disappear in the diff: literals / JSX text / non-ASCII runs on removed lines that do not come back on added lines."""
    diff = git(repo, "diff", "-U0", rng, "--", *globs)
    minus, plus = set(), set()
    for line in diff.splitlines():
        if line[:1] not in "+-" or line[:3] in ("+++", "---"):
            continue
        (minus if line[0] == "-" else plus).update(_texts(line[1:]))
    return sorted(minus - plus)


def used_in_automation(strings):
    hits = {}
    files = [f for d in ("pages", "suites") for f in (WORK / d).rglob("*.py") if not re.search(r"(_text|_runtime)\.py$|^make_", f.name)]
    texts = {f: f.read_text(encoding="utf-8", errors="replace").splitlines() for f in files}
    for s in strings:
        for f, lines in texts.items():
            for i, l in enumerate(lines, 1):
                if s in l:
                    hits.setdefault(s, []).append(f"{f.relative_to(WORK).as_posix()}:{i}")
    return hits


def sysmap_sections(paths):
    docs = {p: p.read_text(encoding="utf-8").splitlines() for p in sorted(SYSMAP.glob("*.md"))} if SYSMAP.is_dir() else {}
    out = {}
    for path in paths:
        parts = path.split("/")
        keys = {path, "/".join(parts[-2:])} | ({"/".join(parts[:-1])} if len(parts) > 2 else set())
        for doc, lines in docs.items():
            head = ""
            for l in lines:
                if l.startswith("#"):
                    head = l.strip("# ").strip()
                if any(k and k in l for k in keys):
                    v = f"{doc.name} › {head[:70]}"
                    if v not in out.setdefault(path, []):
                        out[path].append(v)
    return out


def _in_file(c, wanted):
    name = Path(c["file"]).name
    return "*" in wanted or any((w.isdigit() and f"_{int(w):02d}_" in name) or (not w.isdigit() and w.lower() in name.lower()) for w in wanted)


def analyse(ranges, files=None):
    from ..core.config import cfg
    if not cfg.use_source:
        sys.exit("source.use = false in qa.config.json (black-box mode): impact analysis needs the app source; not available")
    m = json.load(open(MAP, encoding="utf-8")) if MAP.exists() else {"repos": {}, "rules": []}
    cat = cat_mod.load()
    if not cat:
        sys.exit("no catalog: run `python -m qakit.tools.catalog build` first")
    cases = {i: c for i, c in cat["cases"].items() if c["status"] != "deprecated"}
    repos = {k: str((WORK / v).resolve()) if not Path(v).is_absolute() else v for k, v in m.get("repos", {}).items()}
    ch = []
    ranges = {k: norm_range(v) for k, v in (ranges or {}).items()}
    for repo, rng in ranges.items():
        if repo not in repos:
            sys.exit(f"unknown repo '{repo}': add it to test_suite/source_map.json repos")
        ch += [(repo, *x) for x in changed(repos[repo], rng)]
    for f in files or []:
        repo, _, path = f.partition(":")
        ch.append((repo, "M", path.replace("\\", "/"), None))
    ignored, unmapped, mapped, hits, levels, broad = [], [], [], {}, {}, []
    for repo, st, path, old in ch:
        if any(fnmatch.fnmatch(path, g) for g in m.get("ignore", [])):
            ignored.append(f"{repo}:{path}")
            continue
        rules = [r for r in m["rules"] if r["repo"] == repo and (fnmatch.fnmatch(path, r["pattern"]) or (old and fnmatch.fnmatch(old, r["pattern"])))]
        if not rules:
            unmapped.append(f"{repo}:{path} ({st})")
            continue
        mapped.append(dict(repo=repo, status=st, path=path, areas=sorted({r.get("area", "") for r in rules}), map=sorted({r.get("map", "") for r in rules})))
        broad_done = False
        for r in rules:
            why = f"{repo}:{path} → {r.get('area', '')}"
            if r["level"] == "broad":
                if not broad_done:
                    broad.append(why)
                broad_done = True
                continue
            for i, c in cases.items():
                if not _in_file(c, r["files"]):
                    continue
                if r["level"] == "cross" and not any(t.lower() in (c["cls"] + " " + c["typ"]).lower() for t in r.get("tags", [])):
                    continue
                hits.setdefault(i, [])
                if why not in hits[i]:
                    hits[i].append(why)
                if levels.get(i) != "direct":
                    levels[i] = r["level"]
    ui_globs = m.get("ui_globs", ["*.ts", "*.tsx", "*.js", "*.jsx", "*.vue", "*.svelte", "*.html", "*.cshtml", "*.razor", "*.json", "*.resx", "*.properties", "*.po", "*.yml", "*.yaml"])
    rm = sorted({s for repo, rng in ranges.items() for s in removed_strings(repos[repo], rng, ui_globs)})
    broken = used_in_automation(rm)
    stale = sysmap_sections([x["path"] for x in mapped] + [u.split(":", 1)[1].split(" (")[0] for u in unmapped])
    ids = sorted(hits)
    auto = [i for i in ids if cases[i].get("automation")]
    return dict(ranges=ranges, catalog_build=cat["build"], changed=len(ch), mapped=mapped, unmapped=unmapped, ignored=ignored, broad=sorted(set(broad)),
                ids={i: dict(level=levels[i], file=Path(cases[i]["file"]).name, title=cases[i]["title"][:60], automated=bool(cases[i].get("automation")),
                             reasons=hits[i][:5]) for i in ids},
                automated=auto, not_automated=[i for i in ids if i not in auto],
                removed_ui_strings=rm, automation_using_removed_strings=broken, system_map_to_refresh=stale)


def render(r):
    direct = [i for i, v in r["ids"].items() if v["level"] == "direct"]
    cross = [i for i, v in r["ids"].items() if v["level"] == "cross"]
    L = ["# Impact of the source change (automatic first cut)", "",
         f"Ranges: {', '.join(f'{k}=`{v}`' for k, v in r['ranges'].items()) or '-'} · catalog {r['catalog_build']} · {r['changed']} changed file(s) "
         f"({len(r['mapped'])} mapped, {len(r['unmapped'])} unmapped, {len(r['ignored'])} ignored)", "",
         f"**Candidate IDs: {len(r['ids'])}** (direct {len(direct)} · cross by tag {len(cross)}) — automated {len(r['automated'])}, not {len(r['not_automated'])}.",
         "Rough cut from the mapping table; the agent narrows it by reading the diff.", "",
         "## Changed files → areas", "", "| File | Status | Area | system_map |", "|---|---|---|---|"]
    L += [f"| {m['repo']}:{m['path']} | {m['status']} | {'; '.join(m['areas'])} | {', '.join(m['map'])} |" for m in r["mapped"]]
    if r["unmapped"]:
        L += ["", "## Unmapped (agent reads the diff, then adds a rule to test_suite/source_map.json)", *[f"- {u}" for u in r["unmapped"]]]
    if r["broad"]:
        L += ["", "## Broad impact (warning only, no ID added)", *[f"- {b}" for b in r["broad"]]]
    by = {}
    for i, v in r["ids"].items():
        by.setdefault(v["file"], []).append(i)
    L += ["", "## Candidate IDs per workbook", ""]
    for f, ids in sorted(by.items()):
        a = [i for i in ids if r["ids"][i]["automated"]]
        L.append(f"- {f}: {len(ids)} ID (automated {len(a)}) — {', '.join(ids[:25])}{' …' if len(ids) > 25 else ''}")
    if r["automation_using_removed_strings"]:
        L += ["", "## Test code using UI strings removed / changed by the diff (likely broken → automate)", ""]
        for s, locs in r["automation_using_removed_strings"].items():
            L.append(f"- 「{s}」: {', '.join(locs[:6])}{' …' if len(locs) > 6 else ''}")
    elif r["removed_ui_strings"]:
        L += ["", f"UI strings removed / changed: {len(r['removed_ui_strings'])}; none is used by the test code."]
    if r["system_map_to_refresh"]:
        L += ["", "## docs/system_map sections to refresh", *[f"- {p}: {'; '.join(v[:4])}" for p, v in r["system_map_to_refresh"].items()]]
    L += ["", "## Suggested command (after narrowing)", "", f"- `/system-test plan {','.join(r['automated'][:60])}`" if r["automated"] else "- (no automated ID)"]
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--range", action="append", default=[], help="<repo>=<git range>, repeatable")
    ap.add_argument("--files", nargs="*", help="explicit changed files: <repo>:<path>")
    ap.add_argument("--out", help="folder for impact.json + impact.md")
    a = ap.parse_args()
    ranges = dict(x.split("=", 1) for x in a.range)
    if not (ranges or a.files):
        ap.error("give --range <repo>=<range> and/or --files")
    r = analyse(ranges, a.files)
    md = render(r)
    print(md)
    if a.out:
        d = Path(a.out)
        d.mkdir(parents=True, exist_ok=True)
        (d / "impact.json").write_text(json.dumps(r, ensure_ascii=False, indent=1), encoding="utf-8")
        (d / "impact.md").write_text(md, encoding="utf-8")
        print("\nwritten", d / "impact.json", d / "impact.md")


if __name__ == "__main__":
    main()
