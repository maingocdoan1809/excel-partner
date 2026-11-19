"""ANALYZE: what the assistant must work out for every case before anything runs, checked by a script so a half-done analysis cannot pass.

analysis.json = {schema_version, cases: [ {
    id, profile,
    verdict        ready | needs_user | blocked
    clarity        {score 1-5, notes}                    is the case understandable as written
    sufficiency    {score 1-5, missing: [text]}          does it carry the data to be run
    refs           [{text, status: ok | missing | asked | waived, note}]     one per file / record the case points at
    data_needs     [{what, fallback}]                    data the environment may lack, and the step that creates it (ensure / reach)
    edge_cases     [text]                                what can differ between environments; or `edge_cases_none_reason`
    depends_on     [case id]                             cases whose end state this one starts from
    reuse          text                                  what it takes over from them (optional)
    steps          [{n, action, expected}]               the strategy for the case, not the one-line description in the workbook
    timeout_s      int 5-120 (optional)                  how long one wait may last on this case's pages; default 8. A slow screen, a heavy report
    questions / reason                                   needs_user: what to ask; blocked: why
} ]}

`ready` needs clarity >= 3, sufficiency >= 3 and no ref left `missing`, and may depend only on other `ready` cases."""
import hashlib
import json
from pathlib import Path

from . import SCHEMA_VERSION, events, intake, memory, refs
from .store import SchemaError

VERDICTS = ("ready", "needs_user", "blocked")
REF_STATUSES = ("ok", "missing", "asked", "waived")
MIN_TIMEOUT, MAX_TIMEOUT = 5, 120
MIN_WORDS = 12          # characters a written justification must have to say anything


def _says(value, minimum=MIN_WORDS):
    return len(str(value or "").strip()) >= minimum


def template(cases_doc, workspace, session=None):
    """A skeleton to fill in: every case with its steps and the references intake found, judgement fields left empty. With a `session` it also carries what the user said
    and shared since intake (answers, files they pasted by path), and a reference that is missing is looked for again among those files."""
    profiles = [p["name"] for p in workspace["profiles"]]
    stem = memory.stem_of(cases_doc["source"]["name"])
    folders = [Path(cases_doc["source"]["file"]).parent] + [Path(d) for d in workspace.get("reference_dirs", [])]
    items = []
    for case in cases_doc["cases"]:
        entries = []
        for r in case["refs"]:
            path = r.get("path", "") if r["status"] == "found" else (refs.lookup(session, r["text"], folders) if session else "")
            entries.append({"kind": "file", "text": r["text"], "status": "ok" if path else "missing", "path": path, "note": path})
        items.append({
            "id": case["id"], "previous": memory.recall(workspace["name"], stem, case["id"]), "profile": profiles[0] if len(profiles) == 1 else "", "verdict": "",
            "clarity": {"score": 0, "notes": ""}, "sufficiency": {"score": 0, "missing": []},
            "refs": entries,
            "vague": case.get("vague", []),
            "vague_resolved": [{"phrase": v["phrase"], "concrete": "", "source": ""} for v in case.get("vague", [])],
            "data_needs": [], "data_considered": "", "edge_cases": [], "depends_on": list(case["case_refs"]), "reuse": "",
            "steps": [{"n": i, "action": line, "expected": ""} for i, line in enumerate(case["steps"], 1)],
            "_title": case["title"], "_expected": case["expected"],
        })
    out = {"schema_version": SCHEMA_VERSION, "cases": items}
    if session:
        out["_user_answers"] = [{"case": e.get("case"), "answer": e["answer"]} for e in events.read(session.dir) if e["kind"] == "answered"]
        out["_shared_files"] = [{"name": i["name"], "path": i["path"], "kind": i["kind"], "case": i.get("case")} for i in refs.items(session)]
    return out


def _score(value):
    return isinstance(value, int) and not isinstance(value, bool) and 1 <= value <= 5


def validate(analysis, cases_doc, workspace):
    """Raises SchemaError listing every defect."""
    problems = []
    wanted = [c["id"] for c in cases_doc["cases"]]
    given = [c.get("id") for c in analysis.get("cases", []) if isinstance(c, dict)]
    problems += [f"case {i}: not analysed" for i in wanted if i not in given]
    problems += [f"case {i}: not in the testcase file" for i in dict.fromkeys(given) if i not in wanted]
    problems += [f"case {i}: analysed twice" for i in dict.fromkeys(given) if given.count(i) > 1 and i in wanted]
    by_id = {c["id"]: c for c in analysis.get("cases", []) if isinstance(c, dict) and c.get("id") in wanted}
    profiles = [p["name"] for p in workspace["profiles"]]
    info = {c["id"]: c for c in cases_doc["cases"]}
    for ident, c in by_id.items():
        at = f"case {ident}"
        verdict = c.get("verdict")
        if verdict not in VERDICTS:
            problems.append(f"{at}: verdict must be one of {', '.join(VERDICTS)}")
        clarity, sufficiency = c.get("clarity") or {}, c.get("sufficiency") or {}
        if not _score(clarity.get("score")):
            problems.append(f"{at}: clarity.score must be 1-5")
        if not _score(sufficiency.get("score")):
            problems.append(f"{at}: sufficiency.score must be 1-5")
        if c.get("profile") not in profiles:
            problems.append(f"{at}: profile must be one of {', '.join(profiles)}")
        refs = c.get("refs") if isinstance(c.get("refs"), list) else None
        if refs is None:
            problems.append(f"{at}: refs must be a list (empty when the case points at nothing)")
        else:
            for ref in refs:
                if not isinstance(ref, dict) or ref.get("status") not in REF_STATUSES:
                    problems.append(f"{at}: each ref needs a status of {', '.join(REF_STATUSES)}")
                    continue
                label = ref.get("text") or "a reference"
                where = str(ref.get("path") or ref.get("note") or "").strip()
                if ref["status"] == "ok" and (ref.get("kind", "file") == "file") and not (where and Path(where).exists()):
                    problems.append(f"{at}: reference '{label}' is marked ok but there is no file at '{where}'. Give its real path (the user can paste one: `oqa ref add <path>`), or ask for it")
                if ref["status"] == "waived" and not _says(ref.get("note")):
                    problems.append(f"{at}: reference '{label}' is waived: say in its note why the case can be run without it")
        for need in c.get("data_needs") or []:
            if not isinstance(need, dict) or not str(need.get("what", "")).strip() or not str(need.get("fallback", "")).strip():
                problems.append(f"{at}: each data need needs 'what' and a 'fallback' (the step that creates the data)")
        if not c.get("edge_cases") and not str(c.get("edge_cases_none_reason", "")).strip():
            problems.append(f"{at}: list edge_cases, or say why there are none in edge_cases_none_reason")
        steps = c.get("steps")
        if not isinstance(steps, list) or not steps:
            problems.append(f"{at}: steps is empty")
        else:
            for position, step in enumerate(steps, 1):
                if not isinstance(step, dict) or step.get("n") != position or not str(step.get("action", "")).strip() or not str(step.get("expected", "")).strip():
                    problems.append(f"{at}: step {position} needs n={position}, an action and an expected result")
        if "timeout_s" in c and not (isinstance(c["timeout_s"], int) and not isinstance(c["timeout_s"], bool) and MIN_TIMEOUT <= c["timeout_s"] <= MAX_TIMEOUT):
            problems.append(f"{at}: timeout_s must be a whole number of seconds from {MIN_TIMEOUT} to {MAX_TIMEOUT}")
        for dep in c.get("depends_on") or []:
            if dep not in wanted:
                problems.append(f"{at}: depends_on '{dep}' is not a case of this file")
            elif dep == ident:
                problems.append(f"{at}: depends on itself")
            elif dep in by_id and by_id[dep].get("verdict") != "ready" and verdict == "ready":
                problems.append(f"{at}: is ready but depends on {dep}, which is not ready")
        if verdict == "ready":
            if _score(clarity.get("score")) and clarity["score"] < 3:
                problems.append(f"{at}: clarity {clarity['score']} is too low for 'ready'")
            if _score(sufficiency.get("score")) and sufficiency["score"] < 3:
                problems.append(f"{at}: sufficiency {sufficiency['score']} is too low for 'ready'")
            if refs and any(isinstance(r, dict) and r.get("status") in ("missing", "asked") for r in refs):
                problems.append(f"{at}: a reference is still missing or only asked for; ask the user (verdict needs_user) or mark it waived with the reason")
            if sufficiency.get("missing") and not c.get("data_needs"):
                problems.append(f"{at}: sufficiency.missing lists what the case lacks, yet the case is ready with no data_needs: either each missing item has a data_need with the step that creates it, or ask the user")
            problems += _vague_problems(at, c, info.get(ident, {}), refs or [])
        if verdict == "needs_user":
            questions = c.get("questions") or []
            if not questions:
                problems.append(f"{at}: needs_user must list the questions to ask")
            elif any(not _says(q, 10) for q in questions):
                problems.append(f"{at}: every question must be a full sentence the user can answer without opening the workbook (what is unclear, and the options you see)")
            if not _says(c.get("data_considered")):
                problems.append(f"{at}: needs_user needs data_considered: what you tried to settle yourself first (data you could create, a file you could read, the sitemap) and why only the user can answer")
        if verdict == "blocked":
            if not str(c.get("reason", "")).strip():
                problems.append(f"{at}: blocked must say why")
            if not _says(c.get("data_considered")):
                problems.append(f"{at}: blocked needs data_considered: why the data it lacks cannot be created through the UI or an import (a blocked case is the last resort)")
    if not problems and _cycle(by_id):
        problems.append("depends_on has a cycle: " + " -> ".join(_cycle(by_id)))
    if problems:
        raise SchemaError(problems, "analysis")
    return analysis


def _vague_problems(at, c, case, refs):
    """A case is ready only when nothing in it still has to be guessed: every phrase intake flagged ("as designed", "works correctly", no expected result) has been turned into
    something that can be checked, from a file the user gave, from their answer, or (for what cannot be measured) from the workbook or the product; and no step's `expected`
    says something vague again."""
    problems = []
    resolved = {str(r.get("phrase", "")).strip().lower(): r for r in c.get("vague_resolved") or [] if isinstance(r, dict)}
    ok_refs = [str(r.get("text", "")).strip().lower() for r in refs if isinstance(r, dict) and r.get("status") == "ok"]
    for flag in case.get("vague") or []:
        phrase = flag["phrase"]
        item = resolved.get(phrase.strip().lower())
        if not item:
            problems.append(f"{at}: \"{phrase}\" ({flag['where']}) is not concrete. Add it to vague_resolved with what will be checked and where that comes from, or ask the user (needs_user)")
            continue
        concrete, source = str(item.get("concrete", "")).strip(), str(item.get("source", "")).strip()
        if not _says(concrete, 15) or intake.find_vague({"x": concrete}):
            problems.append(f"{at}: \"{phrase}\": write what is checked as facts (which fields, which labels, in which order, which value), not another general phrase")
        if source.lower().startswith("ref:"):
            named = source[4:].strip().lower()
            if not named or not any(named in text or text in named for text in ok_refs if text):
                problems.append(f"{at}: \"{phrase}\": source '{source}' is not one of this case's references marked ok")
        elif source != "user_answer" and not (flag["kind"] == "unmeasurable" and source in ("workbook", "domain")):
            allowed = "ref:<a reference marked ok> or user_answer" + (" (or workbook / domain for what only needs a reading)" if flag["kind"] == "unmeasurable" else "")
            problems.append(f"{at}: \"{phrase}\": a result judged against a document needs its source: {allowed}. A guess is not a source: ask the user")
    for step in c.get("steps") or []:
        if isinstance(step, dict):
            vague = intake.find_vague({"expected": step.get("expected", "")})
            if vague:
                problems.append(f"{at}: step {step.get('n')} expected is still vague (\"{vague[0]['phrase']}\"): say what the page must show")
    return problems


def _cycle(by_id):
    state, trail = {}, []

    def visit(node):
        state[node] = 1
        trail.append(node)
        for dep in by_id[node].get("depends_on") or []:
            if state.get(dep) == 1:
                return trail[trail.index(dep):] + [dep]
            if dep not in state:
                found = visit(dep)
                if found:
                    return found
        state[node] = 2
        trail.pop()
        return None

    for node in by_id:
        if node not in state:
            found = visit(node)
            if found:
                return found
    return None


def order(analysis, cases_doc):
    """Execution order of the `ready` cases: a case after those it depends on; among those free to go, stay on the same profile as the one just run
    (no re-login), then keep the workbook's order. Returns (order, excluded[{id, verdict, reason}])."""
    position = {c["id"]: i for i, c in enumerate(cases_doc["cases"])}
    by_id = {c["id"]: c for c in analysis["cases"]}
    ready = sorted((i for i, c in by_id.items() if c["verdict"] == "ready"), key=position.get)
    excluded = [{"id": i, "verdict": by_id[i]["verdict"], "reason": "; ".join(by_id[i].get("questions") or []) or by_id[i].get("reason", "")}
                for i in sorted(by_id, key=position.get) if by_id[i]["verdict"] != "ready"]
    left, done, out, last_profile = list(ready), set(), [], None
    while left:
        free = [i for i in left if all(d in done for d in by_id[i].get("depends_on") or [])]
        pick = next((i for i in free if by_id[i]["profile"] == last_profile), free[0])
        left.remove(pick)
        done.add(pick)
        out.append(pick)
        last_profile = by_id[pick]["profile"]
    return out, excluded


def digest(obj):
    return hashlib.sha256(json.dumps(obj, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:16]


def verify_strategy(strategy_doc, analysis):
    """True when strategy.json is untouched since `strategy()` made it and still matches the analysis it came from."""
    body = {k: v for k, v in strategy_doc.items() if k != "sha"}
    return digest(body) == strategy_doc.get("sha") and strategy_doc.get("analysis_sha") == digest(analysis)


LABELS = {
    "vi": {"title": "Chiến lược test", "run": "Sẽ chạy theo thứ tự", "skip": "Chưa chạy (cần xử lý trước)", "profile": "Profile", "deps": "Phụ thuộc", "reuse": "Tận dụng",
           "data": "Dữ liệu có thể thiếu → cách bù", "edge": "Tình huống biên", "steps": "Các bước", "clarity": "Rõ ràng", "suff": "Đủ dữ kiện", "none": "không",
           "concrete": "Điều mơ hồ đã được cụ thể hoá thành"},
    "en": {"title": "Test strategy", "run": "Will run in this order", "skip": "Not run yet (needs attention)", "profile": "Profile", "deps": "Depends on", "reuse": "Reuses",
           "data": "Data that may be missing → fallback", "edge": "Edge cases", "steps": "Steps", "clarity": "Clarity", "suff": "Sufficiency", "none": "none",
           "concrete": "Vague wording made concrete"},
}


def strategy(analysis, cases_doc, workspace):
    """(strategy dict, markdown). The digest in the dict is what `confirm` checks, so the user approves exactly what will run."""
    ordered, excluded = order(analysis, cases_doc)
    by_id, info = {c["id"]: c for c in analysis["cases"]}, {c["id"]: c for c in cases_doc["cases"]}
    data = {"schema_version": SCHEMA_VERSION, "order": ordered, "excluded": excluded, "analysis_sha": digest(analysis)}
    data["sha"] = digest(data)
    t = LABELS.get(workspace.get("language"), LABELS["en"])
    lines = [f"# {t['title']}: {cases_doc['source']['name']}", "", f"## {t['run']} ({len(ordered)})", ""]
    for n, ident in enumerate(ordered, 1):
        c = by_id[ident]
        lines += [f"### {n}. {ident}: {info[ident]['title']}", "",
                  f"- {t['profile']}: {c['profile']} · {t['clarity']} {c['clarity']['score']}/5 · {t['suff']} {c['sufficiency']['score']}/5",
                  f"- {t['deps']}: {', '.join(c.get('depends_on') or []) or t['none']}" + (f" · {t['reuse']}: {c['reuse']}" if c.get("reuse") else "")]
        if c.get("data_needs"):
            lines.append(f"- {t['data']}: " + "; ".join(f"{d['what']} → {d['fallback']}" for d in c["data_needs"]))
        resolved = [v for v in c.get("vague_resolved") or [] if v.get("concrete")]
        if resolved:
            lines.append(f"- {t['concrete']}: " + "; ".join(f"\"{v['phrase']}\" → {v['concrete']} [{v['source']}]" for v in resolved))
        lines.append(f"- {t['edge']}: " + ("; ".join(c.get("edge_cases") or []) or c.get("edge_cases_none_reason", "")))
        lines += [f"- {t['steps']}:"] + [f"  {s['n']}. {s['action']} → {s['expected']}" for s in c["steps"]] + [""]
    if excluded:
        lines += [f"## {t['skip']} ({len(excluded)})", ""] + [f"- {e['id']} [{e['verdict']}]: {e['reason']}" for e in excluded] + [""]
    return data, "\n".join(lines)
