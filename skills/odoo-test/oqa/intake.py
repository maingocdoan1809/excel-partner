"""INTAKE: read a testcase workbook into cases.json, and do the part of the analysis a script can do without judgement: split steps, find the files
a case points at (and whether they exist), and find the other cases it mentions. The assistant judges the rest in ANALYZE.

Header names are matched through synonyms (en / ja / vi). The sheet and the header row are detected: the first sheet whose header row (within the
first 15 rows) has an ID column and an expected-result or action column."""
import hashlib
import re
from pathlib import Path

from . import SCHEMA_VERSION
from .store import SchemaError

SYN = {
    "no": ["No.", "No", "ID", "TC ID", "TC No", "Test Case ID", "Case ID", "テストケースID", "ケースID", "テストID", "番号", "ケース番号", "Mã TC", "Mã testcase"],
    "cls": ["分類", "カテゴリ", "Category", "Module", "Feature", "Chức năng", "Phân loại"],
    "typ": ["タイプ", "種別", "Type", "Test Type", "Loại"],
    "view": ["テスト観点", "テスト項目", "テスト名", "試験項目", "Test Case Name", "Test Case", "Title", "Summary", "Objective", "目的",
             "Góc độ kiểm thử", "Quan điểm kiểm thử", "Tên testcase", "Mục đích"],
    "role": ["対象ロール/環境", "ロール", "対象ロール", "実施ロール", "Role", "Roles", "Environment", "Role / Environment", "Role / Môi trường áp dụng", "Vai trò", "Môi trường"],
    "pre": ["事前準備", "前提条件", "事前条件", "Precondition", "Preconditions", "Pre-condition", "Chuẩn bị trước", "Điều kiện tiên quyết", "Tiền điều kiện"],
    "act": ["アクション", "手順", "操作", "操作手順", "Steps", "Action", "Actions", "Test Steps", "Procedure", "Hành động", "Thao tác", "Các bước"],
    "exp": ["期待結果", "期待値", "確認内容", "Expected Result", "Expected Results", "Expected", "Kết quả mong đợi"],
    "note": ["備考", "Notes", "Note", "Remark", "Remarks", "Ghi chú"],
}
FILE_EXT = r"(?:xlsx|xlsm|xls|csv|pdf|docx?|pptx?|png|jpe?g|gif|txt|zip|json|xml|html?)"
QUOTED_FILE = re.compile(r"[\"'`“‘「『]([^\"'`”’」』\n]+?\." + FILE_EXT + r")[\"'`”’」』]", re.I)
BARE_FILE = re.compile(r"(?<![\w/\\.:-])((?:[A-Za-z]:)?[\w\-.()/\\]+\." + FILE_EXT + r")\b", re.I)
NUMBERING = re.compile(r"^\s*(?:\d+[.)\-:：]|[-*・•])\s*")


def _norm(text):
    return re.sub(r"\s+", "", str(text or "")).lower()


_INDEX = {key: {_norm(n) for n in names} for key, names in SYN.items()}


def _match(cell):
    whole = _norm(cell)
    if not whole:
        return None
    for key, names in _INDEX.items():
        if whole in names:
            return key
    for part in re.split(r"[/／\n|]+", str(cell or "")):          # bilingual header: "ケース番号/ID"
        piece = _norm(part)
        if piece and piece != whole:
            for key, names in _INDEX.items():
                if piece in names:
                    return key
    return None


def read_workbook(path):
    """(sheet title, header row, [raw case dicts]). Raises SchemaError when no testcase table is found."""
    import openpyxl
    book = openpyxl.load_workbook(path, data_only=True, read_only=False)
    for sheet in book:
        for header in range(1, 16):
            cols = {}
            for col in range(1, sheet.max_column + 1):
                key = _match(sheet.cell(header, col).value)
                if key and key not in cols:
                    cols[key] = col
            if "no" in cols and ("exp" in cols or "act" in cols):
                rows = []
                for row in range(header + 1, sheet.max_row + 1):
                    ident = sheet.cell(row, cols["no"]).value
                    if ident in (None, ""):
                        continue
                    item = {k: ("" if sheet.cell(row, c).value is None else str(sheet.cell(row, c).value).strip()) for k, c in cols.items()}
                    item["no"] = str(ident).strip()
                    item["_row"] = row
                    rows.append(item)
                return sheet.title, header, rows
    raise SchemaError("no testcase table found: the header row needs an ID column and an expected-result or action column", Path(path).name)


def find_files(text, folders):
    """Every file the text names -> [{text, status: found|missing, path}]. Looked for as given, then in each folder, then by name below the folders."""
    names = []
    for pattern in (QUOTED_FILE, BARE_FILE):
        for match in pattern.finditer(text):
            if match.group(1) not in names:
                names.append(match.group(1))
    out = []
    for name in names:
        found = None
        candidate = Path(name)
        if candidate.is_absolute() and candidate.exists():
            found = candidate
        for folder in folders:
            if found:
                break
            if (folder / name).exists():
                found = folder / name
        for folder in folders:
            if found:
                break
            hit = next(iter(folder.rglob(Path(name).name)), None) if folder.is_dir() else None
            if hit:
                found = hit
        out.append({"kind": "file", "text": name, "status": "found" if found else "missing", "path": str(found) if found else ""})
    return out


def build(path, reference_dirs=()):
    """cases.json content for a testcase workbook."""
    path = Path(path)
    if not path.is_file():
        raise SchemaError(f"file not found: {path}", "testcase file")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise SchemaError("only .xlsx / .xlsm testcase files are read", path.name)
    sheet, header, rows = read_workbook(path)
    folders = [path.parent] + [Path(d) for d in reference_dirs]
    ids = [r["no"] for r in rows]
    warnings = [f"ID '{i}' appears {ids.count(i)} times; the first occurrence is used" for i in dict.fromkeys(ids) if ids.count(i) > 1]
    named = [i for i in dict.fromkeys(ids) if re.search(r"\D", i)]          # a bare number would match any digit in a sentence
    cases, seen = [], set()
    for row in rows:
        if row["no"] in seen:
            continue
        seen.add(row["no"])
        text = "\n".join(row.get(k, "") for k in ("pre", "act", "exp", "note"))
        lines = [NUMBERING.sub("", line).strip() for line in row.get("act", "").split("\n") if line.strip()]
        cases.append({
            "id": row["no"], "title": row.get("view", ""), "category": row.get("cls", ""), "type": row.get("typ", ""), "role": row.get("role", ""),
            "pre": row.get("pre", ""), "steps": lines, "expected": row.get("exp", ""), "note": row.get("note", ""), "sheet": sheet, "row": row["_row"],
            "refs": find_files(text, folders),
            "case_refs": [other for other in named if other != row["no"] and re.search(r"(?<![\w-])" + re.escape(other) + r"(?![\w-])", text)],
        })
    return {"schema_version": SCHEMA_VERSION,
            "source": {"file": str(path.resolve()), "name": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "sheet": sheet, "header_row": header},
            "warnings": warnings, "cases": cases}
