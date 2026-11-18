"""RESOURCES: the files a case needs before it can run (a spreadsheet to import, a picture or a document to attach), prepared BEFORE the user approves the strategy.

The assistant declares them in the analysis (`resources` of a case, `uses` of a step). Each has one of three modes, and the user sees which on the dashboard:
    provided    the user's own file (found next to the workbook, or uploaded on the dashboard)
    described   made from what the user said it should contain
    guessed     made from the assistant's own idea of what the case needs; the user may change it

    sessions/<id>/resources.json   {schema_version, resources: {case: {name: entry}}}
    sessions/<id>/fixtures/<case>/<name>

entry = {name, mode, purpose, description, spec (generated files), source (display name), sha256, bytes, facts, edited}.
`edited` means the user changed it on the dashboard: a later analysis never overwrites that. The user's approval covers `digest_of`, so what they saw is what runs.
Generated files are rebuilt from `spec`, never copied from a path, so a case script never depends on where a file happened to be."""
import csv
import hashlib
import io
import re
import shutil
import textwrap
from dataclasses import dataclass, field
from pathlib import Path

from . import SCHEMA_VERSION
from .store import SchemaError, digest, locked, read_json, write_json

FILE = "resources.json"
MODES = ("provided", "described", "guessed")
GENERATED = ("xlsx", "csv", "png", "jpg", "pdf", "txt")
NAME = re.compile(r"^[A-Za-z0-9_][\w .()-]{0,80}\.[A-Za-z0-9]{1,8}$")
COLOR = re.compile(r"^#[0-9a-fA-F]{6}$")
MAX_ROWS, MAX_COLUMNS, MAX_UPLOAD, MAX_TEXT = 5000, 100, 8 * 1024 * 1024, 200_000
PREVIEW_ROWS = 30
SCALARS = (str, int, float, bool, type(None))


def kind_of(name):
    """`xlsx`, `csv`, `png`, `jpg`, ... from a file name (`jpeg` counts as `jpg`)."""
    ext = name.rsplit(".", 1)[-1].lower()
    return "jpg" if ext == "jpeg" else ext


def fixtures_dir(session_dir, case):
    return Path(session_dir) / "fixtures" / (re.sub(r"[^\w.-]+", "_", case) or "case")


def file_of(session_dir, case, name):
    return fixtures_dir(session_dir, case) / name


def sha_of(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


# --- making the file from a spec -----------------------------------------------------------------
def _fail(*problems):
    raise SchemaError(list(problems), "resource")


def _text(spec, key, default="", limit=200):
    value = spec.get(key, default)
    if not isinstance(value, str) or len(value) > limit:
        _fail(f"{key} must be text of at most {limit} characters")
    return value


def _number(spec, key, default, low, high):
    value = spec.get(key, default)
    if not isinstance(value, int) or isinstance(value, bool) or not low <= value <= high:
        _fail(f"{key} must be a whole number from {low} to {high}")
    return value


def table(spec):
    """(columns, rows) of a spreadsheet spec: {columns: [...], rows: [[...]], count: N}. A text cell may contain `{i}` (the row number, from 1).
    With `count`, the given rows are repeated to make that many (a list longer than one page needs no 150 lines of spec)."""
    columns, rows = spec.get("columns"), spec.get("rows", [])
    if not isinstance(columns, list) or not 1 <= len(columns) <= MAX_COLUMNS or not all(isinstance(c, str) and c.strip() for c in columns):
        _fail(f"columns must be a list of 1 to {MAX_COLUMNS} column titles")
    if len(set(columns)) != len(columns):
        _fail("column titles must be different from each other")
    if not isinstance(rows, list) or not all(isinstance(r, list) and len(r) == len(columns) and all(isinstance(x, SCALARS) for x in r) for r in rows):
        _fail(f"rows must be a list of rows with {len(columns)} values each (text, number, true/false or null)")
    count = spec.get("count")
    if count is not None:
        count = _number(spec, "count", None, 1, MAX_ROWS)
        if not rows:
            _fail("count needs at least one row to repeat")
        rows = [rows[i % len(rows)] for i in range(count)]
    if len(rows) > MAX_ROWS:
        _fail(f"at most {MAX_ROWS} rows")
    return columns, [[x.replace("{i}", str(n)) if isinstance(x, str) else x for x in row] for n, row in enumerate(rows, 1)]


def _spreadsheet(kind, spec):
    columns, rows = table(spec)
    if kind == "csv":
        delimiter, encoding = spec.get("delimiter", ","), spec.get("encoding", "utf-8")
        if delimiter not in (",", ";", "\t") or encoding not in ("utf-8", "utf-8-sig"):
            _fail("delimiter must be , ; or a tab, and encoding utf-8 or utf-8-sig")
        out = io.StringIO()
        writer = csv.writer(out, delimiter=delimiter, lineterminator="\n")
        writer.writerow(columns)
        writer.writerows(rows)
        return out.getvalue().encode(encoding)
    from openpyxl import Workbook
    sheet_name = _text(spec, "sheet", "Sheet1", 31)
    if re.search(r"[\[\]:*?/\\]", sheet_name) or not sheet_name.strip():
        _fail("sheet is a sheet name without [ ] : * ? / \\")
    book = Workbook()
    sheet = book.active
    sheet.title = sheet_name
    sheet.append(columns)
    for row in rows:
        sheet.append(row)
    out = io.BytesIO()
    book.save(out)
    return out.getvalue()


def _font(size):
    from PIL import ImageFont
    try:
        return ImageFont.load_default(size=size)
    except TypeError:                          # an old Pillow has one fixed size
        return ImageFont.load_default()


def _picture(kind, spec):
    from PIL import Image, ImageDraw
    width, height = _number(spec, "width", 640, 16, 4000), _number(spec, "height", 480, 16, 4000)
    color, label = spec.get("color", "#4a90d9"), _text(spec, "text", "")
    if not isinstance(color, str) or not COLOR.match(color):
        _fail("color must look like #4a90d9")
    image = Image.new("RGB", (width, height), color)
    if label:
        draw, font = ImageDraw.Draw(image), _font(max(12, min(height // 8, 48)))
        box = draw.textbbox((0, 0), label, font=font)
        draw.text(((width - (box[2] - box[0])) / 2 - box[0], (height - (box[3] - box[1])) / 2 - box[1]), label, fill="#ffffff", font=font)
    out = io.BytesIO()
    image.save(out, "PNG" if kind == "png" else "JPEG", **({} if kind == "png" else {"quality": 90}))
    return out.getvalue()


def _document(kind, spec):
    from PIL import Image, ImageDraw
    pages, label = _number(spec, "pages", 1, 1, 20), _text(spec, "text", "", 3000)
    font, sheets = _font(14), []
    for number in range(1, pages + 1):
        sheet = Image.new("RGB", (595, 842), "white")
        draw = ImageDraw.Draw(sheet)
        lines = textwrap.wrap(label, 70) or [""]
        for row, line in enumerate(lines[:45] + [f"- {number}/{pages} -"]):
            draw.text((40, 40 + row * 18), line, fill="black", font=font)
        sheets.append(sheet)
    out = io.BytesIO()
    sheets[0].save(out, "PDF", save_all=True, append_images=sheets[1:], resolution=72)
    return out.getvalue()


def _plain(kind, spec):
    return _text(spec, "text", "", MAX_TEXT).encode("utf-8")


def render(kind, spec):
    """The bytes of a file of `kind` made from `spec`. Raises SchemaError listing what is wrong with the spec."""
    if kind not in GENERATED:
        _fail(f"a .{kind} file cannot be made from a description: it has to come from the user (mode provided). Makeable: {', '.join(GENERATED)}")
    if not isinstance(spec, dict):
        _fail("spec must be an object")
    return {"xlsx": _spreadsheet, "csv": _spreadsheet, "png": _picture, "jpg": _picture, "pdf": _document, "txt": _plain}[kind](kind, spec)


# --- what a file contains ------------------------------------------------------------------------
def _scan(path, kind, prefix="", keep=PREVIEW_ROWS):
    """{columns, rows (the first `keep`, as text), total (data rows), prefixed (some cell starts with the test prefix)} of a csv / xlsx file."""
    columns, kept, total, prefixed = [], [], 0, False
    if kind == "csv":
        text = Path(path).read_text(encoding="utf-8-sig", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        source, book = csv.reader(io.StringIO(text), dialect), None
    else:
        from openpyxl import load_workbook
        book = load_workbook(path, read_only=True, data_only=True)
        source = book.active.iter_rows(values_only=True)
    try:
        for number, row in enumerate(source):
            cells = ["" if x is None else str(x) for x in row]
            if number == 0:
                columns = [c for c in cells]
                while columns and not columns[-1]:
                    columns.pop()
                continue
            if not any(cells):
                continue
            total += 1
            if len(kept) < keep:
                kept.append(cells[:len(columns) or None])
            prefixed = prefixed or bool(prefix) and any(c.startswith(prefix) for c in cells)
    finally:
        if book is not None:
            book.close()
    return {"columns": columns, "rows": kept, "total": total, "prefixed": prefixed}


def facts_of(path, prefix=""):
    """What the dashboard and the script may say about a file: its size and, by kind, rows and columns / picture size / lines. Never raises."""
    path = Path(path)
    kind = kind_of(path.name)
    out = {"bytes": path.stat().st_size}
    try:
        if kind in ("csv", "xlsx"):
            scan = _scan(path, kind, prefix, keep=0)
            out.update(rows=scan["total"], columns=scan["columns"][:MAX_COLUMNS])
            if prefix:
                out["prefix_missing"] = bool(scan["total"]) and not scan["prefixed"]
        elif kind in ("png", "jpg"):
            from PIL import Image
            with Image.open(path) as image:
                out.update(width=image.width, height=image.height)
        elif kind == "txt":
            out["lines"] = len(path.read_text(encoding="utf-8", errors="replace").splitlines())
    except Exception as error:                 # a damaged file the user uploaded: say so, do not stop the dashboard
        out["unreadable"] = f"{type(error).__name__}: {str(error)[:120]}"
    return out


def preview(session_dir, case, name, limit=PREVIEW_ROWS):
    """The first rows of a spreadsheet, the start of a text file, or the picture's size: what the user looks at before approving."""
    entry = load(session_dir)["resources"].get(case, {}).get(name)
    path = file_of(session_dir, case, name)
    if not entry or not path.is_file():
        raise SchemaError(f"{case} has no file {name} yet", "resource")
    kind = kind_of(name)
    if kind in ("csv", "xlsx"):
        try:
            scan = _scan(path, kind, keep=limit)
        except Exception as error:
            raise SchemaError(f"{name} cannot be read as a {kind} file ({type(error).__name__}: {str(error)[:100]})", "resource")
        return {"kind": kind, "columns": scan["columns"], "rows": scan["rows"], "total": scan["total"]}
    if kind == "txt":
        return {"kind": kind, "text": path.read_text(encoding="utf-8", errors="replace")[:2000]}
    return {"kind": kind, **{k: v for k, v in entry.get("facts", {}).items() if k in ("width", "height", "bytes")}}


# --- the declarations in the analysis ------------------------------------------------------------
def declared(case_analysis):
    return [r for r in case_analysis.get("resources") or [] if isinstance(r, dict) and r.get("name")]


def problems(c, at):
    """Every defect in a case's `resources` and in its steps' `uses`, as sentences for analysis.validate."""
    out, items = [], c.get("resources")
    items = [] if items is None else items
    if not isinstance(items, list):
        return [f"{at}: resources must be a list"]
    names = []
    for position, r in enumerate(items, 1):
        if not isinstance(r, dict) or not isinstance(r.get("name"), str) or not NAME.match(r["name"]):
            out.append(f"{at}: resource {position} needs a `name` that is a file name with an extension (letters, digits, space and . _ - ( ))")
            continue
        name, label = r["name"], f"{at}: resource {r['name']}"
        if name in names:
            out.append(f"{label} is declared twice")
        names.append(name)
        mode = r.get("mode")
        if mode not in MODES:
            out.append(f"{label}: mode must be one of {', '.join(MODES)}")
        if not str(r.get("purpose", "")).strip():
            out.append(f"{label}: say what it is for (purpose)")
        if mode in ("described", "guessed"):
            try:
                render(kind_of(name), r.get("spec"))
            except SchemaError as error:
                out += [f"{label}: {p}" for p in error.problems]
        if mode == "described" and not str(r.get("description", "")).strip():
            out.append(f"{label}: a described file needs `description`, the user's own words")
        if mode == "provided" and r.get("source") and not Path(str(r["source"])).is_file():
            out.append(f"{label}: source file not found: {r['source']} (leave `source` empty when the user will upload it on the dashboard)")
    used = set()
    for step in c.get("steps") if isinstance(c.get("steps"), list) else []:
        uses = step.get("uses", []) if isinstance(step, dict) else []
        if not isinstance(uses, list) or not all(isinstance(u, str) for u in uses):
            out.append(f"{at}: step {step.get('n') if isinstance(step, dict) else '?'}: uses must be a list of resource names")
            continue
        for u in uses:
            if u not in names:
                out.append(f"{at}: step {step.get('n')} uses '{u}', which is not a resource of this case")
        used.update(uses)
    out += [f"{at}: resource {n} is not used by any step: name it in the `uses` of the step that needs it" for n in names if n not in used]
    return out


# --- resources.json ------------------------------------------------------------------------------
def load(session_dir):
    data = read_json(Path(session_dir) / FILE, None) or {"schema_version": SCHEMA_VERSION}
    data.setdefault("resources", {})
    return data


def _stamp(entry, target, prefix=""):
    entry.update(sha256=sha_of(target), bytes=Path(target).stat().st_size, facts=facts_of(target, prefix))
    return entry


def _write(target, content):
    Path(target).parent.mkdir(parents=True, exist_ok=True)
    Path(target).write_bytes(content)


def _prepare(session_dir, case, d, prior, prefix):
    """The entry for one declaration: the user's own edit when there is one, else the declared file made (or copied) now."""
    name, mode = d["name"], d["mode"]
    target = file_of(session_dir, case, name)
    if prior and prior.get("edited"):
        return {**prior, "purpose": d["purpose"].strip()}
    entry = {"name": name, "mode": mode, "purpose": d["purpose"].strip(), "description": str(d.get("description", "")).strip(), "edited": False}
    if mode == "provided":
        source = str(d.get("source") or "")
        if source:
            entry["source"] = Path(source).name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
            return _stamp(entry, target, prefix)
        target.unlink(missing_ok=True)
        return entry                                                    # waits for the user's file
    spec = d["spec"]
    entry["spec"] = spec
    if prior and prior.get("spec") == spec and prior.get("mode") != "provided" and target.is_file() and sha_of(target) == prior.get("sha256"):
        return {**entry, **{k: prior[k] for k in ("sha256", "bytes", "facts")}}        # same spec as last time: keep the file the user may already have looked at
    _write(target, render(kind_of(name), spec))
    return _stamp(entry, target, prefix)


def sync(session_dir, analysis_doc, prefix=""):
    """Make resources.json and the files match the declarations of an analysis that was just accepted. What the user changed on the dashboard stays.
    Returns {resources, guessed, needs_file}."""
    path = Path(session_dir) / FILE
    with locked(path):
        data = load(session_dir)
        old, new = data["resources"], {}
        for c in analysis_doc["cases"]:
            for d in declared(c):
                new.setdefault(c["id"], {})[d["name"]] = _prepare(session_dir, c["id"], d, old.get(c["id"], {}).get(d["name"]), prefix)
        for case, items in old.items():
            for name in items:
                if name not in new.get(case, {}):
                    file_of(session_dir, case, name).unlink(missing_ok=True)
        data["resources"] = new
        write_json(path, data)
    entries = [e for mine in new.values() for e in mine.values()]
    return {"resources": len(entries), "guessed": sum(e["mode"] == "guessed" for e in entries), "needs_file": sum(not e.get("sha256") for e in entries)}


def edit(session_dir, case, name, spec=None, upload=None, filename="", prefix=""):
    """The user's change to one resource: a new `spec` (the file is made again, mode becomes described) or an uploaded file (mode becomes provided)."""
    path = Path(session_dir) / FILE
    with locked(path):
        data = load(session_dir)
        entry = data["resources"].get(case, {}).get(name)
        if entry is None:
            raise SchemaError(f"{case} has no resource {name}", "resource")
        target = file_of(session_dir, case, name)
        if spec is not None:
            _write(target, render(kind_of(name), spec))
            entry.update(mode="described", spec=spec, edited=True)
            entry.pop("source", None)
        elif upload is not None:
            if not upload or len(upload) > MAX_UPLOAD:
                raise SchemaError(f"the file must not be empty or larger than {MAX_UPLOAD // (1024 * 1024)} MB", "resource")
            if kind_of(filename or name) != kind_of(name):
                raise SchemaError(f"{name} is a .{kind_of(name)} file: upload a .{kind_of(name)} file", "resource")
            _write(target, upload)
            entry.update(mode="provided", edited=True, source=Path(filename or name).name)
            entry.pop("spec", None)
        else:
            raise SchemaError("nothing to change: give a spec or a file", "resource")
        _stamp(entry, target, prefix)
        write_json(path, data)
    return entry


def status_of(session_dir, case, entry):
    """ready | needs_file (a provided file nobody gave yet) | missing (the file is gone) | changed (the file is not what was recorded)."""
    if not entry.get("sha256"):
        return "needs_file"
    target = file_of(session_dir, case, entry["name"])
    if not target.is_file():
        return "missing"
    return "ready" if sha_of(target) == entry["sha256"] else "changed"


def digest_of(session_dir):
    """What the user approves: for every resource its mode and its spec (or, for a file they gave, its content). Changes when anything they saw changes."""
    items = [[case, name, e.get("mode"), e["spec"] if e.get("spec") is not None else e.get("sha256")]
             for case, mine in sorted(load(session_dir)["resources"].items()) for name, e in sorted(mine.items())]
    return digest(items)


def blocking(session_dir, analysis_doc, order):
    """Why the strategy cannot be approved yet: a resource of a case that will run which has no file."""
    data, by_id, out = load(session_dir)["resources"], {c["id"]: c for c in analysis_doc["cases"]}, []
    for case in order:
        for d in declared(by_id.get(case, {})):
            entry = data.get(case, {}).get(d["name"])
            state = status_of(session_dir, case, entry) if entry else "unprepared"
            if state == "needs_file":
                out.append(f"{case}: {d['name']} needs the user's file: upload it on the dashboard")
            elif state != "ready":
                out.append(f"{case}: {d['name']} is {state}" + (": submit the analysis again" if state == "unprepared" else ""))
    return out


def view(session_dir, case, case_analysis):
    """One record per declared resource of a case, for the dashboard and the report (no paths)."""
    mine, out = load(session_dir)["resources"].get(case, {}), []
    for d in declared(case_analysis):
        entry = mine.get(d["name"])
        facts = (entry or {}).get("facts") or {}
        warnings = (["unreadable"] if facts.get("unreadable") else []) + (["prefix"] if facts.get("prefix_missing") else [])
        out.append({"name": d["name"], "kind": kind_of(d["name"]), "mode": (entry or d)["mode"], "purpose": d["purpose"], "description": (entry or {}).get("description", ""),
                    "status": status_of(session_dir, case, entry) if entry else "unprepared", "edited": bool(entry and entry.get("edited")), "source": (entry or {}).get("source", ""),
                    "spec": (entry or {}).get("spec"), "makeable": kind_of(d["name"]) in GENERATED, "bytes": (entry or {}).get("bytes", 0), "facts": facts, "warnings": warnings,
                    "used_by": [s["n"] for s in case_analysis.get("steps") or [] if d["name"] in (s.get("uses") or [])]})
    return out


MODE_LABELS = {"vi": {"provided": "bạn cung cấp", "described": "theo mô tả của bạn", "guessed": "AI đoán"},
               "en": {"provided": "your file", "described": "from your description", "guessed": "AI's guess"}}
STATUS_LABELS = {"vi": {"needs_file": "chưa có file", "missing": "mất file", "changed": "file đã bị đổi", "unprepared": "chưa chuẩn bị"},
                 "en": {"needs_file": "no file yet", "missing": "file missing", "changed": "file was changed", "unprepared": "not prepared"}}


def describe(item, language="en"):
    """One line about a resource for strategy.md and the terminal: name, who decided its content, what it is for, what is inside."""
    lang = "vi" if language == "vi" else "en"
    facts, bits = item["facts"], []
    if "rows" in facts:
        bits.append(f"{facts['rows']} " + ("dòng" if lang == "vi" else "rows") + f" x {len(facts.get('columns', []))} " + ("cột" if lang == "vi" else "columns"))
    if "width" in facts:
        bits.append(f"{facts['width']}x{facts['height']}")
    if item["status"] != "ready":
        bits.append(STATUS_LABELS[lang].get(item["status"], item["status"]))
    return f"{item['name']} [{MODE_LABELS[lang][item['mode']]}] {item['purpose']}" + (f" ({', '.join(bits)})" if bits else "")


# --- what a script receives ----------------------------------------------------------------------
@dataclass(frozen=True)
class Fixture:
    """A prepared file, as `s.fixture(name)` hands it to a step: `path` is what a file input takes, the rest says what is inside (use it in checks: rows, columns)."""
    name: str
    path: str
    kind: str
    mode: str
    sha256: str
    bytes: int
    facts: dict = field(default_factory=dict)

    @property
    def rows(self):
        return self.facts.get("rows")

    @property
    def columns(self):
        return self.facts.get("columns", [])

    def __str__(self):
        return self.path

    def record(self):
        """What the step event keeps about the file."""
        return {"name": self.name, "kind": self.kind, "mode": self.mode, "sha256": self.sha256[:12], "bytes": self.bytes, **({"rows": self.rows} if self.rows is not None else {})}


def fixture(session_dir, state, case, name):
    """The prepared file `name` of `case`, checked: the user approved exactly these resources and the file is exactly what was recorded."""
    entry = load(session_dir)["resources"].get(case, {}).get(name)
    if entry is None:
        raise SchemaError(f"{case} has no resource {name}: declare it in the analysis", "resource")
    approved = (state.get("confirmation") or {}).get("resources_sha")
    if approved != digest_of(session_dir):
        raise SchemaError("the resources changed after the user approved the strategy: they must approve again", "resource")
    problem = status_of(session_dir, case, entry)
    if problem != "ready":
        raise SchemaError(f"{name} is {problem}: it cannot be used", "resource")
    path = file_of(session_dir, case, name)
    return Fixture(name, str(path), kind_of(name), entry["mode"], entry["sha256"], entry["bytes"], facts_of(path))
