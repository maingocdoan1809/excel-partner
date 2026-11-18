# -*- coding: utf-8 -*-
"""Read ANY testcase workbook into normalised dicts. Header names are matched through synonyms (en / ja / vi + the project's own names in
qa.config.json testcases.columns). The sheet and the header row are auto-detected: the first sheet whose header row (within the first 15
rows) contains an ID column AND an expected-result (or action) column."""
import re
import openpyxl
from ..core.config import cfg

SYN = {
    "no":   ["No.", "No", "ID", "TC ID", "TC No", "Test Case ID", "Case ID", "テストケースID", "ケースID", "テストID", "番号", "ケース番号", "Mã TC", "Mã testcase"],
    "cls":  ["分類", "カテゴリ", "Category", "Module", "Feature", "Chức năng", "Phân loại"],
    "typ":  ["タイプ", "種別", "Type", "Test Type", "Loại"],
    "view": ["テスト観点", "テスト項目", "テスト名", "試験項目", "Test Case Name", "Test Case", "Title", "Summary", "Objective", "目的",
             "Góc độ kiểm thử", "Quan điểm kiểm thử", "Tên testcase", "Mục đích"],
    "role": ["対象ロール/環境", "ロール", "対象ロール", "実施ロール", "Role", "Roles", "Environment", "Role / Environment", "Role / Môi trường áp dụng",
             "Vai trò", "Môi trường"],
    "pre":  ["事前準備", "前提条件", "事前条件", "Precondition", "Preconditions", "Pre-condition", "Chuẩn bị trước", "Điều kiện tiên quyết", "Tiền điều kiện"],
    "act":  ["アクション", "手順", "操作", "操作手順", "Steps", "Action", "Actions", "Test Steps", "Procedure", "Hành động", "Thao tác", "Các bước"],
    "exp":  ["期待結果", "期待値", "確認内容", "Expected Result", "Expected Results", "Expected", "Kết quả mong đợi"],
    "result": ["結果", "Result", "Actual Result", "Kết quả", "Kết quả thực tế"],
    "tester": ["Người thực hiện", "Người test", "実施者", "担当者", "Tester", "Executed by", "Executor"],
    "date": ["Ngày thực hiện", "Ngày test", "Ngày", "実施日", "日付", "Test date", "Execution date", "Executed on", "Date"],
    "note": ["備考", "Notes", "Note", "Remark", "Remarks", "Ghi chú"],
}
for k, extra in (cfg.raw["testcases"].get("columns") or {}).items():       # project header names, e.g. {"exp": ["Kỳ vọng"]}
    SYN.setdefault(k, [])
    SYN[k] = list(extra) + SYN[k]


def _norm(s):
    return re.sub(r"\s+", "", str(s or "")).lower()


_IDX = {k: {_norm(n) for n in names} for k, names in SYN.items()}


def _match(cell):
    c = _norm(cell)
    if not c:
        return None
    for k, names in _IDX.items():
        if c in names:
            return k
    for part in re.split(r"[/／\n|]+", str(cell or "")):        # bilingual header: "ケース番号/ID", "タイプ\nType"
        p = _norm(part)
        if p and p != c:
            for k, names in _IDX.items():
                if p in names:
                    return k
    return None


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb:
        for r in range(1, 16):
            cols = {}
            for c in range(1, ws.max_column + 1):
                k = _match(ws.cell(r, c).value)
                if k and k not in cols:
                    cols[k] = c
            if "no" in cols and ("exp" in cols or "act" in cols):
                cases = []
                for rr in range(r + 1, ws.max_row + 1):
                    v = ws.cell(rr, cols["no"]).value
                    if v in (None, ""):
                        continue
                    d = {k: (ws.cell(rr, c).value if ws.cell(rr, c).value is not None else "") for k, c in cols.items()}
                    d["no"] = str(v).strip() if not isinstance(v, int) else v
                    d["act_lines"] = [x.strip() for x in str(d.get("act", "")).split("\n") if x.strip()]
                    d["_sheet"], d["_row"], d["_header"] = ws.title, rr, r
                    d["_cols"] = {k: str(ws.cell(r, c).value).strip() for k, c in cols.items()}
                    cases.append(d)
                return cases
    raise RuntimeError(f"no testcase table found in {path} (needs an ID column and an expected-result or action column; "
                       f"add the project's header names to qa.config.json testcases.columns)")


def by_no(cases):
    return {c["no"]: c for c in cases}
