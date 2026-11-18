"""REPORT: the Excel evidence of a test run, written for the person who has to read it.

  report/Evidence_<testcase file>_<session>.xlsx   ONE workbook for the whole run: a "Tổng hợp" sheet (every case with its result), then one sheet per case
  report/SUMMARY.md                                the result in a few lines

One file, rebuilt in place each time: no pile of per-case workbooks.

A case sheet says: what the testcase asked (from the workbook), how the assistant understood and planned it, what happened, and every step with the picture
of the screen, red-framed, taken for it. Built from view.case_records, the same record the dashboard shows."""
import math
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from . import view
from .state import Session

LABELS = {
    "vi": {
        "title": "Báo cáo evidence kiểm thử", "summary": "Tổng hợp", "file": "File testcase", "system": "Hệ thống", "session": "Phiên test", "started": "Bắt đầu", "finished": "Kết thúc",
        "totals": "Kết quả chung", "no": "No.", "category": "Phân loại", "name": "Test Case Name", "pre": "Điều kiện tiên quyết", "src_steps": "Các bước (file gốc)",
        "src_exp": "Kết quả mong đợi (file gốc)", "profile": "Profile", "result": "Kết quả", "runs": "Số lần chạy", "checks": "Số kiểm tra", "remark": "Nhận xét",
        "s1": "1. Testcase yêu cầu gì (theo file gốc)", "s2": "2. AI hiểu và lên kế hoạch thế nào", "s3": "3. Kết quả", "s4": "4. Từng bước và ảnh evidence (lần chạy cuối)", "s5": "5. Các lần chạy",
        "note": "Ghi chú trong file", "row": "Dòng trong file", "deps": "Tiếp nối từ", "rating": "Đánh giá đề bài", "clarity": "Rõ ràng", "suff": "Đủ dữ kiện", "edge": "Tình huống biên đã tính",
        "data": "Dữ liệu có thể thiếu → cách bù", "files": "File tài nguyên", "file_used": "File dùng", "mode_provided": "người dùng cung cấp",
        "mode_described": "AI sinh theo mô tả của người dùng", "mode_guessed": "AI tự đoán", "edited": "người dùng đã sửa", "asked": "Câu hỏi AI cần hỏi bạn", "hints": "Gợi ý của bạn", "repairs": "AI phân tích lại và sửa",
        "r_after": "Sau lần {n}", "r_cause": "nguyên nhân", "r_evidence": "căn cứ", "r_change": "đã sửa", "r_reread": "hiểu lại case", "verdict": "Kết luận", "step": "Bước", "action": "Thao tác",
        "expected": "Kết quả mong đợi", "actual": "Thực tế", "nchecks": "Kiểm tra", "evidence": "Ảnh evidence", "attempt": "Lần", "failed": "Bước lỗi", "reason": "Lý do",
        "notrun": "Chưa chạy", "none": "(không có)", "case_sheet": "Mở sheet",
        "ok": "Đạt: {n} bước đều đã chạy, có đối chiếu với kết quả mong đợi và có ảnh evidence.", "fail": "Không đạt{at}: {why}", "pending": "Chờ bạn quyết định: {why}",
        "blocked": "Không thực hiện được: {why}", "notrun_text": "Chưa chạy trong lần test này{why}", "retry": " Đã phải chạy {n} lần (các lần trước lỗi, xem mục 5).",
    },
    "en": {
        "title": "Test evidence report", "summary": "Summary", "file": "Testcase file", "system": "System", "session": "Test session", "started": "Started", "finished": "Finished",
        "totals": "Overall", "no": "No.", "category": "Category", "name": "Test Case Name", "pre": "Precondition", "src_steps": "Steps (source file)",
        "src_exp": "Expected result (source file)", "profile": "Profile", "result": "Result", "runs": "Runs", "checks": "Checks", "remark": "Remark",
        "s1": "1. What the testcase asks (from the source file)", "s2": "2. How the assistant understood and planned it", "s3": "3. Result", "s4": "4. Each step and its evidence picture (last run)", "s5": "5. Runs",
        "note": "Note in the file", "row": "Row in the file", "deps": "Continues from", "rating": "Rating of the case", "clarity": "Clarity", "suff": "Sufficiency", "edge": "Edge cases considered",
        "data": "Data that may be missing → fallback", "files": "Resource files", "file_used": "File used", "mode_provided": "provided by the user",
        "mode_described": "made from the user's description", "mode_guessed": "the assistant's guess", "edited": "edited by the user", "asked": "Questions for you", "hints": "Your hints", "repairs": "How the assistant re-analysed and repaired",
        "r_after": "After run {n}", "r_cause": "cause", "r_evidence": "evidence", "r_change": "changed", "r_reread": "now reads the case as", "verdict": "Conclusion", "step": "Step", "action": "Action",
        "expected": "Expected result", "actual": "Actual", "nchecks": "Checks", "evidence": "Evidence picture", "attempt": "Run", "failed": "Failed step", "reason": "Reason",
        "notrun": "Not run", "none": "(none)", "case_sheet": "Open sheet",
        "ok": "Passed: all {n} steps ran, each was checked against its expected result and has an evidence picture.", "fail": "Failed{at}: {why}", "pending": "Waiting for your decision: {why}",
        "blocked": "Could not be done: {why}", "notrun_text": "Not run in this test{why}", "retry": " It took {n} runs (the earlier ones failed, see section 5).",
    },
}
COLORS = {"PASS": "C6EFCE", "FAIL": "FFC7CE", "PENDING": "FFEB9C", "BLOCKED": "D9D9D9"}
NOT_RUN_COLOR = "EDEDED"
IMAGE_WIDTH = 440
WIDTHS = {"A": 22, "B": 40, "C": 40, "D": 40, "E": 10, "F": 11, "G": 64}
TEXT_CHARS = 200            # roughly how many characters fit on one line of a merged B:G cell


def shown(status, t):
    return status if status in COLORS else t["notrun"]


def paint(cell, label):
    from openpyxl.styles import PatternFill
    cell.fill = PatternFill("solid", fgColor=COLORS.get(label, NOT_RUN_COLOR))


def height_for(text, chars):
    return 15 * sum(max(1, math.ceil(len(part) / chars)) for part in str(text or "").split("\n")) + 4


def conclusion(record, last, t):
    status = record["status"]
    why = record["reason"] or (last["reason"] if last else "")
    if status == "PASS":
        return t["ok"].format(n=len(last["steps"]) if last else 0) + (t["retry"].format(n=len(record["attempts"])) if len(record["attempts"]) > 1 else "")
    if status == "FAIL":
        at = f" ({t['step'].lower()} {last['failed_step']})" if last and last.get("failed_step") else ""
        return t["fail"].format(at=at, why=why or "-")
    if status == "PENDING":
        return t["pending"].format(why=why or "-")
    if status == "BLOCKED":
        return t["blocked"].format(why=why or "-")
    return t["notrun_text"].format(why=f": {why}" if why else ".")


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:28] or "case"
    candidate, n = base, 2
    while candidate.lower() in used:
        candidate, n = f"{base[:25]}_{n}", n + 1
    used.add(candidate.lower())
    return candidate


def write_case(sheet, session: Session, record, t):
    """One case, top to bottom, on `sheet`."""
    from openpyxl.drawing.image import Image
    from openpyxl.styles import Alignment, Font, PatternFill
    for column, width in WIDTHS.items():
        sheet.column_dimensions[column].width = width
    wrap = Alignment(wrap_text=True, vertical="top")
    state = session.state()
    source, label = record["source"], shown(record["status"], t)
    last = record["attempts"][-1] if record["attempts"] else None
    row = 1

    def line(key, value, color=None):
        nonlocal row
        sheet.cell(row, 1, t[key]).font = Font(bold=True)
        sheet.cell(row, 1).alignment = wrap
        sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        cell = sheet.cell(row, 2, value if value not in (None, "") else t["none"])
        cell.alignment = wrap
        if color:
            paint(cell, color)
        sheet.row_dimensions[row].height = height_for(value, TEXT_CHARS)
        row += 1

    def section(key):
        nonlocal row
        row += 1
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        cell = sheet.cell(row, 1, t[key])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F3A4F")
        sheet.row_dimensions[row].height = 20
        row += 1

    sheet.merge_cells("A1:G1")
    head = sheet.cell(1, 1, f"{record['id']}  ·  {record['title']}  ·  {label}")
    head.font = Font(bold=True, size=14)
    paint(head, label)
    sheet.row_dimensions[1].height = 26
    row = 2
    section("s1")
    line("no", record["id"])
    line("name", record["title"])
    line("category", " / ".join(x for x in (source["category"], source["type"]) if x))
    line("pre", source["pre"])
    line("src_steps", "\n".join(f"{i}. {s}" for i, s in enumerate(source["steps"], 1)))
    line("src_exp", source["expected"])
    line("note", source["note"])
    line("row", record["row"])
    section("s2")
    line("profile", record["profile"])
    line("deps", ", ".join(record["depends_on"]) + (f" ({record['reuse']})" if record["reuse"] else "") if record["depends_on"] else "")
    line("rating", f"{t['clarity']} {record['clarity']}/5 · {t['suff']} {record['sufficiency']}/5" if record["clarity"] else "")
    line("edge", "\n".join(f"• {e}" for e in record["edge_cases"]) or record["edge_none"])
    line("data", "\n".join(f"• {d['what']} → {d['fallback']}" for d in record["data_needs"]))
    if record["resources"]:
        line("files", "\n".join(f"• {r['name']} [{t['mode_' + r['mode']]}{', ' + t['edited'] if r['edited'] else ''}]: {r['purpose']}" + (f" ({r['facts']['rows']} rows)" if "rows" in r["facts"] else "") for r in record["resources"]))
    if record["questions"]:
        line("asked", "\n".join(f"• {q}" for q in record["questions"]))
    section("s3")
    line("result", label, color=label)
    line("runs", len(record["attempts"]))
    if record["hints"]:
        line("hints", "\n".join(f"• {h['text']}" for h in record["hints"]))
    if record["repairs"]:
        line("repairs", "\n".join(f"• {t['r_after'].format(n=r['attempt'])}: {t['r_cause']}: {r['cause']}; {t['r_evidence']}: {r['evidence']}; {t['r_change']}: {r['change']}"
                                  + (f"; {t['r_reread']}: {r['reanalysis']}" if r.get("reanalysis") else "") for r in record["repairs"]))
    line("verdict", conclusion(record, last, t))
    section("s4")
    for column, key in enumerate(("step", "action", "expected", "actual", "nchecks", "result", "evidence"), 1):
        cell = sheet.cell(row, column, t[key])
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
    row += 1
    steps = last["steps"] if last else [{"n": p["n"], "caption": p["caption"], "expected": p["expected"], "status": "PENDING", "actual": "", "images": [], "checks": None} for p in record["planned"]]
    for step in steps:
        step_label = shown(step["status"], t)
        actual = step["actual"] + "".join(f"\n{t['file_used']}: {r['name']} [{t['mode_' + r['mode']]}]" for r in step.get("resources") or [])
        for column, value in enumerate((step["n"], step["caption"], step["expected"], actual, step.get("checks") if step.get("checks") is not None else "", step_label), 1):
            sheet.cell(row, column, value).alignment = wrap
        paint(sheet.cell(row, 6), step_label)
        height = max(height_for(step["caption"], 42), height_for(step["expected"], 42), height_for(actual, 42))
        picture = session.dir / step["images"][0]["annotated"] if step["images"] else None
        if picture and picture.exists():
            image = Image(str(picture))
            image.height, image.width = round(image.height * IMAGE_WIDTH / image.width), IMAGE_WIDTH
            sheet.add_image(image, f"G{row}")
            height = max(height, image.height * 0.75 + 6)
        sheet.row_dimensions[row].height = height
        row += 1
    section("s5")
    for column, key in enumerate(("attempt", "result", "failed", "reason"), 1):
        cell = sheet.cell(row, column, t[key])
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
    row += 1
    for attempt in record["attempts"]:
        status = shown(attempt["status"], t) if attempt["status"] in COLORS else attempt["status"]
        for column, value in enumerate((attempt["n"], status, attempt["failed_step"] or "", attempt["reason"]), 1):
            sheet.cell(row, column, value).alignment = wrap
        paint(sheet.cell(row, 2), status)
        sheet.row_dimensions[row].height = height_for(attempt["reason"], 42)
        row += 1
    sheet.freeze_panes = "A2"
    sheet.sheet_view.zoomScale = 90


def write_summary(sheet, session: Session, ws, records, t, sheet_names):
    from openpyxl.styles import Alignment, Font, PatternFill
    state = session.state()
    log = view.events.read(session.dir)
    counts = {}
    for record in records:
        label = shown(record["status"], t)
        counts[label] = counts.get(label, 0) + 1
    widths = {"A": 10, "B": 14, "C": 34, "D": 32, "E": 44, "F": 44, "G": 12, "H": 12, "I": 8, "J": 9, "K": 54}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.merge_cells("A1:K1")
    sheet["A1"] = f"{t['title']}: {Path(state['file']).name}"
    sheet["A1"].font = Font(bold=True, size=15)
    version = (state.get("connect") or {}).get("version")
    facts = [("file", state["file"]), ("system", f"{ws['url']}" + (f" · db {ws['db']}" if ws.get("db") else "") + (f" · Odoo {version}" if version else "")),
             ("session", f"{ws['name']} / {session.id}"), ("started", state["created"]), ("finished", log[-1]["ts"] if log else ""),
             ("totals", " · ".join(f"{k} {v}" for k, v in sorted(counts.items())) + f"  ({len(records)})")]
    for offset, (key, value) in enumerate(facts, 2):
        sheet.cell(offset, 1, t[key]).font = Font(bold=True)
        sheet.merge_cells(start_row=offset, start_column=3, end_row=offset, end_column=11)
        sheet.cell(offset, 3, value)
    top = len(facts) + 3
    headers = ("no", "category", "name", "pre", "src_steps", "src_exp", "profile", "result", "runs", "checks", "remark")
    for column, key in enumerate(headers, 1):
        cell = sheet.cell(top, column, t[key])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F3A4F")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    wrap = Alignment(wrap_text=True, vertical="top")
    for offset, record in enumerate(records, 1):
        row = top + offset
        last = record["attempts"][-1] if record["attempts"] else None
        label = shown(record["status"], t)
        total_checks = sum(s.get("checks") or 0 for s in last["steps"]) if last else 0
        values = (record["id"], " / ".join(x for x in (record["source"]["category"], record["source"]["type"]) if x), record["title"], record["source"]["pre"],
                  "\n".join(f"{i}. {s}" for i, s in enumerate(record["source"]["steps"], 1)), record["source"]["expected"], record["profile"], label,
                  len(record["attempts"]) or "", total_checks or "", conclusion(record, last, t))
        for column, value in enumerate(values, 1):
            sheet.cell(row, column, value).alignment = wrap
        paint(sheet.cell(row, 8), label)
        link = sheet.cell(row, 1)
        link.hyperlink = f"#'{sheet_names[record['id']]}'!A1"
        link.font = Font(color="0563C1", underline="single")
        sheet.row_dimensions[row].height = max(height_for(values[4], 46), height_for(values[10], 56), height_for(values[5], 46))
    sheet.freeze_panes = sheet.cell(top + 1, 4)
    sheet.sheet_view.zoomScale = 90


def build(session: Session, ws):
    """Write the evidence workbook and SUMMARY.md. Returns {folder, main, files, counts}."""
    from openpyxl import Workbook
    t = LABELS.get(ws.get("language"), LABELS["en"])
    records = sorted(view.case_records(session), key=lambda r: r["row"])
    folder = session.dir / "report"
    folder.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w.-]+", "_", Path(session.state()["file"]).stem)
    main = folder / f"Evidence_{stem}_{session.id}.xlsx"
    for old in folder.glob("*.xlsx"):
        if old != main and not old.name.startswith("~$"):           # `~$...` is Excel's lock file for a workbook that is open
            try:
                old.unlink(missing_ok=True)                         # one current workbook per session; also clears the per-case files older versions wrote
            except OSError:
                pass                                                # open in Excel right now: it goes at the next rebuild
    shutil.rmtree(folder / "cases", ignore_errors=True)
    book = Workbook()
    summary = book.active
    summary.title = t["summary"]
    used = {t["summary"].lower()}
    names = {r["id"]: safe_sheet_name(r["id"], used) for r in records}
    for record in records:
        write_case(book.create_sheet(names[record["id"]]), session, record, t)
    write_summary(summary, session, ws, records, t, names)
    try:
        book.save(main)
    except PermissionError:                                         # the person has the workbook open in Excel: write beside it instead of failing
        main = main.with_name(main.stem + time.strftime("_%H%M%S") + ".xlsx")
        book.save(main)
    totals = view.counts(records)
    lines = [f"# {Path(session.state()['file']).name}", "", ", ".join(f"{k} {v}" for k, v in sorted(totals.items())), ""]
    lines += [f"- {r['id']} [{r['status']}] {r['title']}" + (f": {r['reason']}" if r["reason"] else "") for r in records]
    (folder / "SUMMARY.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"folder": str(folder), "main": str(main), "files": [str(main), str(folder / "SUMMARY.md")], "counts": totals}


def finish(session: Session, ws):
    """The report of a finished session: build it, then close the session."""
    session.require("REPORT")
    result = build(session, ws)
    session.done(result["counts"])
    return result


def open_file(path):
    """Show the workbook to the person: their spreadsheet program opens it. Skipped when OQA_NO_OPEN is set (tests, servers)."""
    if os.environ.get("OQA_NO_OPEN"):
        return False
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except OSError:
        return False
