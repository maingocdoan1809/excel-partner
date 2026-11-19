"""Reference files and folders the user points at: a design (BD), a spec, an import file, a folder of attachments.

The user pastes a PATH (in the dashboard, in `oqa ref add`, or inside an answer): nothing is uploaded or copied, the assistant reads the file where it is
(`oqa ref show`). A path written into an answer is picked up by itself (`register_from_text`), so the user never has to learn a command.

refs.json (in the session folder) = {schema_version, items:[{path, name, kind: file|folder, by: user|answer, case, added}]}"""
import html
import os
import re
import time
import zipfile
from pathlib import Path

from . import SCHEMA_VERSION, events
from .store import SchemaError, locked, read_json, write_json

FILE = "refs.json"
SECRET = re.compile(r"[/\\]sessions[/\\][^/\\]+[/\\]auth|storagestate|[/\\]credentials", re.I)
WINDOWS_START = re.compile(r"(?:^|(?<=[\s\"'`“‘「『(\[,;:=]))(?:[A-Za-z]:[\\/]|\\\\[\w.$-]+[\\/]|~[\\/]|%\w+%[\\/])")
UNIX_START = re.compile(r"(?:^|(?<=[\s\"'`“‘「『(\[,;:=]))/[\w.~-][^\r\n]*")
QUOTED = re.compile(r"[\"'`“‘「『]([^\"'`”’」』\r\n]{3,}?)[\"'`”’」』]")
TRAILING = " \t.,;:!?)]}>\"'`”’」』。、"
TEXT_SUFFIXES = (".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".log", ".yml", ".yaml", ".py", ".sql")
LOOK_AT = (".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")


class RefError(Exception):
    pass


def _file(session):
    return Path(session.dir) / FILE


def load(session):
    return read_json(_file(session)) or {"schema_version": SCHEMA_VERSION, "items": []}


def items(session):
    return load(session)["items"]


def _key(path):
    return os.path.normcase(str(path))


def _clean(raw):
    text = str(raw or "").strip().strip("\"'`“”‘’「」『』").strip()
    return os.path.expandvars(text)


def refuse_secret(path):
    if SECRET.search(str(path)):
        raise RefError("Signed-in cookies and passwords are never read: that path is not allowed.")


def add(session, raw, by="user", case=None):
    """Remember a path the user gave. Returns the item. Raises RefError when it does not exist or must not be read."""
    text = _clean(raw)
    if not text:
        raise RefError("Paste the path of a file or folder first.")
    target = Path(text).expanduser()
    if not target.exists():
        raise RefError(f"Not found: {text}. Check the path (it is read on this computer, nothing is uploaded).")
    target = target.resolve()
    refuse_secret(target)
    item = {"path": str(target), "name": target.name or str(target), "kind": "folder" if target.is_dir() else "file", "by": by, "case": case or None,
            "added": time.strftime("%Y-%m-%d %H:%M:%S")}
    with locked(_file(session)):
        data = load(session)
        for old in data["items"]:
            if _key(old["path"]) == _key(item["path"]):
                if case and not old.get("case"):
                    old["case"] = case
                    write_json(_file(session), data)
                return old
        data["items"].append(item)
        write_json(_file(session), data)
    return item


def extract_paths(text):
    """Paths that exist on this computer, written anywhere in a piece of text (quoted or not, with spaces or not)."""
    found, seen = [], set()

    def keep(candidate):
        candidate = _clean(candidate)
        try:
            if not os.path.exists(candidate):
                candidate = candidate.rstrip(TRAILING)
            if candidate and os.path.exists(candidate) and _key(candidate) not in seen:
                seen.add(_key(candidate))
                found.append(candidate)
                return True
        except (OSError, ValueError):
            pass
        return False

    def longest(start_text):
        """`start_text` runs to the end of the line: the path is its longest prefix that exists, cut at a space."""
        words = start_text.split(" ")
        for count in range(len(words), 0, -1):
            if keep(" ".join(words[:count])):
                return True
        return False

    for line in str(text or "").splitlines():
        for match in QUOTED.finditer(line):
            if re.match(r"(?:[A-Za-z]:[\\/]|\\\\|~[\\/]|/|%)", match.group(1).strip()):
                keep(match.group(1))
        for pattern in (WINDOWS_START, UNIX_START):
            for match in pattern.finditer(line):
                longest(line[match.start():].lstrip())
    return found


def register_from_text(session, text, by="answer", case=None):
    """Every existing path written in `text` becomes a reference. Never raises: a path that cannot be used is skipped."""
    added = []
    for candidate in extract_paths(text):
        try:
            added.append(add(session, candidate, by=by, case=case))
        except RefError:
            continue
    return added


def add_many(session, text, by="user", case=None):
    """One path per line (or paths written inside a sentence). -> (added items, lines that are not a path on this computer)."""
    added, failed = [], []
    for line in str(text or "").splitlines():
        if not line.strip():
            continue
        try:
            added.append(add(session, line, by=by, case=case))
        except RefError:
            found = register_from_text(session, line, by=by, case=case)
            if found:
                added += found
            else:
                failed.append(line.strip())
    return added, failed


def folders(session):
    return [Path(i["path"]) for i in items(session) if i["kind"] == "folder" and Path(i["path"]).is_dir()]


def lookup(session, name, extra_folders=()):
    """The path of a reference by its name or path: registered files, registered folders (searched below), then `extra_folders`. '' when there is none."""
    wanted = _clean(name)
    if not wanted:
        return ""
    direct = Path(wanted).expanduser()
    if direct.is_absolute() and direct.exists():
        return str(direct)
    base = Path(wanted).name.lower()
    for item in items(session):
        if item["kind"] == "file" and (_key(item["path"]) == _key(wanted) or item["name"].lower() == base) and Path(item["path"]).exists():
            return item["path"]
    for folder in list(folders(session)) + [Path(f) for f in extra_folders]:
        if not folder.is_dir():
            continue
        if (folder / wanted).exists():
            return str(folder / wanted)
        hit = next((p for p in folder.rglob("*") if p.name.lower() == base), None)
        if hit:
            return str(hit)
    return ""


def resolver(session, config):
    """name -> path, for `o.ref`: the files the user pointed at, then the workbook's folder and the workspace's reference folders."""
    def resolve(name):
        folders = [Path(session.cases()["source"]["file"]).parent] + [Path(d) for d in config.get("reference_dirs", [])]
        return lookup(session, name, folders)
    return resolve


def listing(session):
    rows = items(session)
    if not rows:
        return "No reference file yet. The user can paste a path in the dashboard (Tệp tham chiếu) or in an answer; you can add one with `oqa ref add <path>`."
    return "\n".join(f"- {i['name']} [{i['kind']}] {i['path']}" + (f" (case {i['case']})" if i.get("case") else "") for i in rows)


# --- reading ------------------------------------------------------------------------------
def media_of(path):
    """Names of the pictures embedded in an office file (xlsx / docx / pptx)."""
    try:
        with zipfile.ZipFile(path) as z:
            return [n for n in z.namelist() if re.match(r"(?:xl|word|ppt)/media/[^/]+$", n)]
    except (zipfile.BadZipFile, OSError):
        return []


def workbook_extras(path):
    """What a testcase workbook holds that intake does not read as text: (sheet names, picture count, comment count)."""
    import openpyxl
    book = openpyxl.load_workbook(path, data_only=True)
    comments = sum(1 for sheet in book for row in sheet.iter_rows() for cell in row if cell.comment)
    return [s.title for s in book if s.max_row > 1 or s.max_column > 1], len(media_of(path)), comments


def _extract_media(path, out_dir):
    names = media_of(path)
    saved = []
    if names and out_dir:
        out = Path(out_dir) / re.sub(r"[^\w.-]+", "_", Path(path).stem)
        out.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path) as z:
            for name in names:
                target = out / Path(name).name
                target.write_bytes(z.read(name))
                saved.append(str(target))
    return names, saved


def _cut(lines, limit):
    out, size = [], 0
    for index, line in enumerate(lines):
        size += len(line) + 1
        if size > limit:
            out.append(f"... {len(lines) - index} more line(s); raise --max, or ask for one --sheet")
            break
        out.append(line)
    return out


def _xlsx(path, limit, sheet):
    import openpyxl
    book = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for ws in book:
        if sheet and ws.title.lower() != sheet.lower():
            lines.append(f"## sheet {ws.title} ({ws.max_row} rows) - skipped, ask with --sheet \"{ws.title}\"")
            continue
        lines.append(f"## sheet {ws.title} ({ws.max_row} rows x {ws.max_column} columns)")
        for row in ws.iter_rows():
            cells = [str(c.value).strip().replace("\n", " / ")[:120] for c in row if c.value not in (None, "")]
            notes = [f"[comment {c.coordinate}: {c.comment.text.strip()[:200]}]" for c in row if c.comment]
            if cells or notes:
                lines.append(f"{row[0].row}: " + " | ".join(cells + notes))
    return lines


def _ooxml_text(path, pattern, parts):
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist() if re.match(parts, n)), key=lambda n: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", n)])
        out = []
        for name in names:
            xml = z.read(name).decode("utf-8", "replace")
            if "word/" in name:
                for paragraph in re.split(r"</w:p>", xml):
                    text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", paragraph, re.S))
                    if text.strip():
                        out.append(html.unescape(text))
            else:
                out.append(f"## {Path(name).stem}")
                out += [html.unescape(t) for t in re.findall(pattern, xml, re.S) if t.strip()]
        return out


def describe(path, limit=6000, sheet=None, out_dir=None):
    """The content of a reference as text a person would copy out of it. Pictures, scans and PDFs cannot be read as text: the answer says which file to open with the Read tool."""
    target = Path(path)
    refuse_secret(target)
    if not target.exists():
        raise RefError(f"Not found: {path}")
    if target.is_dir():
        files = sorted(p for p in target.rglob("*") if p.is_file())
        rows = [f"{p.relative_to(target)} ({p.stat().st_size} bytes)" for p in files[:80]]
        return f"Folder {target}: {len(files)} file(s)\n" + "\n".join(rows) + (f"\n... {len(files) - 80} more" if len(files) > 80 else "")
    suffix = target.suffix.lower()
    head = f"{target.name} ({target.stat().st_size} bytes)"
    if suffix in (".xlsx", ".xlsm"):
        lines = _xlsx(target, limit, sheet)
    elif suffix == ".docx":
        lines = _ooxml_text(target, "", r"word/document\.xml$")
    elif suffix == ".pptx":
        lines = _ooxml_text(target, r"<a:t>(.*?)</a:t>", r"ppt/slides/slide\d+\.xml$")
    elif suffix in TEXT_SUFFIXES:
        lines = target.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    elif suffix in LOOK_AT:
        return f"{head}: a picture or PDF cannot be read as text. Open it with the Read tool: {target}"
    elif suffix == ".zip":
        with zipfile.ZipFile(target) as z:
            return head + "\n" + "\n".join(z.namelist()[:80])
    else:
        return f"{head}: no text reader for {suffix or 'this file'}; open it with the Read tool if it is text or an image."
    out = [head] + _cut(lines, limit)
    names, saved = _extract_media(target, out_dir) if suffix in (".xlsx", ".xlsm", ".docx", ".pptx") else ([], [])
    if names:
        out.append(f"Embedded pictures: {len(names)}" + (f", saved here so you can look at them with the Read tool:\n" + "\n".join(saved) if saved else ""))
    return "\n".join(out)


def announce(session, item):
    """The user shared something outside a question: the assistant is told (`oqa wait`)."""
    events.append(session.dir, "ref_added", path=item["path"], by=item["by"], name=item["name"], what=item["kind"], **({"case": item["case"]} if item.get("case") else {}))


__all__ = ["RefError", "SchemaError", "add", "add_many", "announce", "describe", "extract_paths", "folders", "items", "listing", "lookup", "media_of", "register_from_text", "workbook_extras"]
