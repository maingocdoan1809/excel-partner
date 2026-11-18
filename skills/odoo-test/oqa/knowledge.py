"""knowledge/index.md of a workspace: one line per fact, so `grep` finds it. What the assistant learned about THIS Odoo (a screen's quirk, a stable
selector, the page after login, an error and the fix that worked).

    - [screen:sale.order] [kind:quirk] [tag:save] The Save button only enables after a customer is chosen -> -
"""
import re

from .paths import workspace_dir

KINDS = ("login", "navigation", "locator", "data", "error", "fix", "quirk", "rule")
HEADER = """# Knowledge index

One line per fact. Format:  - [screen:<screen>] [kind:<kind>] [tag:<word>] <fact> -> <file or ->
kind: login | navigation | locator | data | error | fix | quirk | rule
Never write a password, token or cookie here.

"""
SECRETISH = re.compile(r"(password|passwd|mật\s*khẩu|token|cookie|session_id|api[_-]?key|secret)\s*[:=]", re.I)
LINE = re.compile(r"^- ((?:\[[a-z]+:[^\]]+\]\s*)+)(.*)$")


class KnowledgeError(Exception):
    pass


def file_of(ws):
    return workspace_dir(ws) / "knowledge" / "index.md"


def _key(screen, kind, fact):
    return (screen.strip().lower(), kind, re.sub(r"\s+", " ", fact).strip().lower())


def parse(text):
    out = []
    for number, raw in enumerate(text.splitlines(), 1):
        match = LINE.match(raw.strip())
        if not match:
            continue
        labels = {k: v for k, v in re.findall(r"\[([a-z]+):([^\]]+)\]", match.group(1))}
        fact, _, ref = match.group(2).partition(" -> ")
        out.append({"line": number, "screen": labels.get("screen", ""), "kind": labels.get("kind", ""), "tag": labels.get("tag", ""), "fact": fact.strip(), "ref": ref.strip() or "-"})
    return out


def search(ws, screen=None, kind=None, text=None):
    path = file_of(ws)
    items = parse(path.read_text(encoding="utf-8")) if path.exists() else []
    return [i for i in items if (not screen or screen.lower() in i["screen"].lower()) and (not kind or i["kind"] == kind)
            and (not text or text.lower() in (i["fact"] + " " + i["tag"]).lower())]


def related(ws, text, limit=3):
    """Known errors, fixes, quirks and locators that share at least two words with `text`: what a failed step should be compared with first."""
    words = {w for w in re.findall(r"\w{4,}", text.lower())}
    scored = []
    for item in search(ws):
        if item["kind"] in ("error", "fix", "quirk", "locator"):
            common = len(words & set(re.findall(r"\w{4,}", (item["fact"] + " " + item["tag"]).lower())))
            if common >= 2:
                scored.append((common, f"[screen:{item['screen']}] [kind:{item['kind']}] {item['fact']} -> {item['ref']}"))
    return [line for _, line in sorted(scored, key=lambda s: -s[0])[:limit]]


def learn(ws, screen, kind, fact, tag="", ref="-"):
    """Append one fact at the END of the file. Returns 'added' or 'duplicate'. Refuses an unknown kind or anything that looks like a secret."""
    screen, fact, tag, ref = (str(x or "").strip() for x in (screen, fact, tag, ref))
    if kind not in KINDS:
        raise KnowledgeError(f"kind must be one of {', '.join(KINDS)}")
    if not screen or not fact:
        raise KnowledgeError("a fact needs a screen and a sentence")
    if "\n" in fact or "]" in screen or "]" in tag:
        raise KnowledgeError("one line per fact: no line breaks, no ']' in screen or tag")
    if SECRETISH.search(fact):
        raise KnowledgeError("that looks like a password, token or cookie; write what you learned without the secret")
    path = file_of(ws)
    existing = parse(path.read_text(encoding="utf-8")) if path.exists() else []
    if any(_key(i["screen"], i["kind"], i["fact"]) == _key(screen, kind, fact) for i in existing):
        return "duplicate"
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        path.write_text(HEADER, encoding="utf-8")
    with open(path, "a", encoding="utf-8") as handle:
        handle.write(f"- [screen:{screen}] [kind:{kind}]" + (f" [tag:{tag}]" if tag else "") + f" {fact} -> {ref or '-'}\n")
    return "added"
