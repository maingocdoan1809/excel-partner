# -*- coding: utf-8 -*-
"""Evidence workbook builder (any suite, any language).
  python -m qakit.reporting.excel_builder --suite <suite> --spec <xlsx> [--lang vi|ja|en] [--tr suites/<suite>/<lang>_text.py]
                                          [--only-run] [--consolidated]
Output: excel/<spec name>_evidence_<lang>.xlsx (consolidated: ..._<lang>_tong_hop.xlsx). Layout: one summary sheet, the source sheets of
the testcase workbook (translated when `--tr` is given), one sheet per case: result, time, environment, spec text, every step with its
verdict, every image (900 px wide) with its caption.
`tr` (optional, when the testcase / captions are in another language than the workbook): module with SPEC{id:{view,role,pre,act[],exp}},
CLS{}, TYP{}, COVER{}, SHEETS{}, CAP{(id,i):str}, ACT{(id,i):str}, REM{id:str}, NOTE{}. Status colours: PASS green / FAIL red /
BLOCKED yellow / PENDING orange / NOT EXECUTED grey."""
import datetime, os
import openpyxl
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageDraw
from ..core import evidence, results
from ..core.config import cfg
from . import spec as spec_mod

BOLD = Font(bold=True, size=11)
WRAP = Alignment(wrap_text=True, vertical="top")
FILL = {"PASS": "FFE2EFDA", "FAIL": "FFF8CBAD", "BLOCKED": "FFFFF2CC", "PENDING": "FFFCE4D6", "NOT EXECUTED": "FFEDEDED"}

LABELS = {
    "en": dict(result="Result", when="Executed at", env="Environment", url="URL", clstyp="Category / Type", role="Role / environment",
               pre="Precondition", acts="Steps (spec)", exp="Expected result (spec)", total="Steps", step="Step", note="Note", screen="Screen",
               summary="Summary", sum_title="System test evidence", sum_env="Environment",
               sum_note="Verdicts are based on real operations, screens and data reads; nothing is mocked or guessed. BLOCKED = cannot be verified because of environment / permission / data limits.",
               agg="Total", cols=["No.", "Test viewpoint", "Result", "Executed at", "Screen evidence", "Key points"], n_img="{n} images (see the sheet for each step)",
               spec_cnt="{a} spec steps / {b} images", op="Operation", actual="Actual", remark="Remark", dur="Duration (per environment · role)",
               src="Source of the result", this_run="This run ({run})", carried="Result of run {run} ({when}); not in the scope of this run",
               spec_chg="\n※ The spec changed after that run: run it again", spec_unk="\n※ Cannot compare with the spec of that run (no catalog history)",
               never="Not executed (in no run)", cons_note="Consolidated: this run {a} · from earlier runs {b} (spec changed since: {c}) · never run {n}"),
    "vi": dict(result="Kết quả", when="Thời gian thực hiện", env="Môi trường", url="URL", clstyp="Phân loại / Loại", role="Vai trò / môi trường",
               pre="Chuẩn bị trước", acts="Các bước (theo đặc tả)", exp="Kết quả mong đợi (theo đặc tả)", total="Tổng số bước", step="Bước", note="Ghi chú",
               screen="Màn hình", summary="Tổng hợp", sum_title="Minh chứng kiểm thử hệ thống", sum_env="Môi trường",
               sum_note="Kết luận dựa trên thao tác thực tế, hiển thị màn hình và kết quả đọc dữ liệu; không dùng mock / suy đoán. BLOCKED = không kiểm chứng được do giới hạn môi trường / quyền / dữ liệu.",
               agg="Tổng hợp", cols=["No.", "Góc độ kiểm thử", "Kết quả", "Thời gian thực hiện", "Ảnh minh chứng", "Điểm chính của đánh giá"],
               n_img="{n} ảnh (ghi chú từng bước xem trong sheet)", spec_cnt="{a} bước đặc tả / {b} ảnh minh chứng", op="Thao tác", actual="Kết quả thực tế",
               remark="Ghi chú", dur="Thời gian thực thi (theo môi trường · vai trò)",
               src="Nguồn kết quả", this_run="Lần chạy này (run {run})", carried="Kết quả lấy từ run {run} ({when}); không nằm trong phạm vi lần chạy này",
               spec_chg="\n※ Đặc tả đã đổi sau lần chạy đó: cần chạy lại", spec_unk="\n※ Không so được với đặc tả lúc chạy (thiếu lịch sử catalog)",
               never="Chưa chạy (không có trong lần chạy nào)", cons_note="Bản tổng hợp: chạy lần này {a} · lấy từ run trước {b} (trong đó {c} case đặc tả đã đổi) · chưa chạy {n}"),
    "ja": dict(result="結果", when="実施日時", env="実施環境", url="URL", clstyp="分類 / タイプ", role="対象ロール/環境", pre="事前準備", acts="手順(仕様)",
               exp="期待結果(仕様)", total="総ステップ数", step="Step", note="補足", screen="画面", summary="サマリ", sum_title="システムテスト 実施エビデンス",
               sum_env="実施環境", sum_note="判定は実機操作・画面表示・データ読取結果に基づく。モック/推測による判定はしていない。BLOCKED は環境・権限・データの制約により検証不能。",
               agg="集計", cols=["No.", "テスト観点", "結果", "実施日時", "画面エビデンス", "判定の要点"], n_img="{n}枚（各ステップの操作メモはシート参照）",
               spec_cnt="仕様手順 {a} 手順 / 証跡 {b} 枚", op="操作", actual="実際結果", remark="Remark", dur="実行時間（環境・ロール別）",
               src="結果の出典", this_run="今回の実施（run {run}）", carried="過去の実施結果を引用：run {run}（{when}）。今回の実施範囲外",
               spec_chg="\n※ その実施後に仕様が変更された：再実施が必要", spec_unk="\n※ その実施時点の仕様と比較できない（catalog 履歴なし）",
               never="未実施（どの実施にも含まれていない）", cons_note="統合版：今回の実施 {a} 件 / 過去の実施から引用 {b} 件（うち仕様変更後 {c} 件）/ 未実施 {n} 件"),
}


def est_h(text, width=120):
    def w(l):
        return sum(2 if ord(ch) > 0x2E80 else 1.1 for ch in l)
    return max(18, 19 * sum(max(1, -(-int(w(l)) // (width - 30))) for l in str(text).split(chr(10))) + 4)


def step_images(s, i):
    ims = s.get("images")
    if ims:
        return [dict(x, annotated=results.locate(x["annotated"]), tag=x.get("tag") or str(i), note=x.get("note") or s["caption"]) for x in ims if x.get("annotated")]
    return [dict(annotated=results.locate(s["annotated"]), tag=str(i), note=s["caption"])] if s.get("annotated") else []


def _recaption(src, dst, label):
    im = Image.open(src).convert("RGB")
    W, H = im.size
    canvas = Image.new("RGB", (W, H))
    canvas.paste(im.crop((0, 0, W, H - evidence.CAP)), (0, 0))
    dr = ImageDraw.Draw(canvas)
    y0 = H - evidence.CAP
    dr.rectangle([0, y0, W, H], fill=(255, 245, 245))
    dr.line([0, y0, W, y0], fill=(255, 0, 0), width=3)
    sz = 24
    while evidence.text_width(dr, label, sz) > W - 36 and sz > 12:
        sz -= 1
    evidence.draw_text(dr, (18, y0 + evidence.CAP / 2), label, sz, (200, 0, 0))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    canvas.save(dst)


ALIASES = {"PASS": ["OK", "Pass", "PASS", "Đạt", "合格", "○"], "FAIL": ["NG", "Fail", "FAIL", "Không đạt", "不合格", "×"],
           "PENDING": ["Hoãn", "Pending", "PENDING", "保留"], "BLOCKED": ["N/A", "Blocked", "BLOCKED", "対象外", "-"]}


def _allowed(ws, col_letter, row):
    """Values of a list data validation covering <col><row> (the customer's vocabulary, e.g. OK / NG / Hoãn / N/A), else None."""
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1 and dv.formula1.startswith('"') and f"{col_letter}{row}" in dv.sqref:
            return [x.strip() for x in dv.formula1.strip('"').split(",")]
    return None


def _fill_results(wb, RES, report):
    """Fill the result / tester / date columns of the testcase table in the OUTPUT workbook (the source file is never written):
    result = status mapped to the customer's vocabulary (report.result_values, else the list validation of the column, else the status);
    NOT EXECUTED (and cases without a result) stay empty; tester = report.tester; date = the day the case was executed. Notes untouched."""
    from openpyxl.utils import get_column_letter
    mapping = report.get("result_values") or {}
    tester = report.get("tester") or ""
    n = 0
    for ws in wb:
        hdr = cols = None
        for r in range(1, 16):
            cc = {spec_mod._match(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if spec_mod._match(ws.cell(r, c).value)}
            if "no" in cc and ("exp" in cc or "act" in cc):
                hdr, cols = r, cc
                break
        if not hdr or "result" not in cols:
            continue
        for r in range(hdr + 1, ws.max_row + 1):
            no = ws.cell(r, cols["no"]).value
            rec = RES.get(str(no).strip()) if no not in (None, "") else None
            if not rec or rec.get("status") in (None, "NOT EXECUTED"):
                continue
            st = rec["status"]
            allowed = _allowed(ws, get_column_letter(cols["result"]), r)
            val = mapping.get(st)
            if val is None and allowed:
                val = next((a for a in ALIASES.get(st, []) if a in allowed), st)
            ws.cell(r, cols["result"]).value = val if val is not None else st
            if tester and "tester" in cols:
                ws.cell(r, cols["tester"]).value = tester
            if "date" in cols and rec.get("started"):
                ws.cell(r, cols["date"]).value = str(rec["started"])[:10]
            n += 1
    return n


def _translate_sources(wb, tr):
    """Translate the source sheets: free texts through tr.COVER, the testcase table through tr.SPEC (no structural change)."""
    cov = getattr(tr, "COVER", {})
    for ws in wb:
        hdr = cols = None
        for r in range(1, 16):
            cc = {spec_mod._match(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if spec_mod._match(ws.cell(r, c).value)}
            if "no" in cc and ("exp" in cc or "act" in cc):
                hdr, cols = r, cc
                break
        for row in ws.iter_rows(min_row=1, max_row=hdr or ws.max_row):
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in cov:
                    c.value = cov[c.value.strip()]
        if not hdr:
            continue
        for r in range(hdr + 1, ws.max_row + 1):
            sp = getattr(tr, "SPEC", {}).get(ws.cell(r, cols["no"]).value)
            if not sp:
                continue
            for k in ("view", "role", "pre", "exp"):
                if k in cols and sp.get(k):
                    ws.cell(r, cols[k]).value = sp[k]
            if "act" in cols and sp.get("act"):
                ws.cell(r, cols["act"]).value = "\n".join(sp["act"])
            for k, table in (("cls", "CLS"), ("typ", "TYP")):
                if k in cols:
                    v = ws.cell(r, cols[k]).value
                    ws.cell(r, cols[k]).value = getattr(tr, table, {}).get(v, v)
    for ws in wb:
        new = getattr(tr, "SHEETS", {}).get(ws.title)
        if new:
            ws.title = new


def build(spec_path, lang="vi", tr=None, env_text=None, url=None, only_run=False, consolidated=False):
    L = LABELS.get(lang, LABELS["en"])
    only_run = only_run and not consolidated
    RES = results.load()
    if spec_path:
        cases = spec_mod.by_no(spec_mod.read(spec_path))
        if only_run:
            cases = {k: v for k, v in cases.items() if str(k) in RES}
        wb = openpyxl.load_workbook(spec_path)
    else:
        cases = {k: dict(no=k, view=v.get("title", ""), role="", pre="", exp="", cls="", typ="", act_lines=[s["action"] for s in v.get("steps", [])])
                 for k, v in RES.items()}
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    if spec_path:
        _fill_results(wb, RES, dict(cfg.raw.get("report", {}), tester=cfg.raw.get("report", {}).get("tester") or cfg.account))     # before translation: headers are still the original ones
    if tr is not None:
        _translate_sources(wb, tr)
    env_text = env_text or f"{cfg.name} / {cfg.base_url}" + (f" / {cfg.account}" if cfg.account else "")
    url = url or cfg.base_url
    lang_dir = cfg.evidence_lang_dir(lang) / (cfg.suite or "")
    summary = []
    for no in sorted(cases, key=lambda x: (str(type(x)), x)):
        c = cases[no]
        sp = tr.SPEC[no] if tr is not None and no in getattr(tr, "SPEC", {}) else dict(view=c.get("view", ""), role=c.get("role", ""), pre=c.get("pre", ""),
                                                                                      act=c["act_lines"], exp=c.get("exp", ""))
        rec = RES.get(str(no), {})
        status = rec.get("status", "NOT EXECUTED")
        steps = rec.get("steps", [])
        sh = wb.create_sheet((f"No.{no:02d}" if isinstance(no, int) else f"{no}")[:31])
        sh.sheet_format.defaultRowHeight = 13
        sh.column_dimensions["A"].width = 30
        sh.column_dimensions["B"].width = 120
        sh.cell(1, 1, f"No.{no}" if isinstance(no, int) else str(no)).font = Font(bold=True, size=12)
        sh.cell(1, 2, str(sp["view"])).font = Font(size=11)
        cur = [3]

        def row(k, v, fill=None):
            r = cur[0]
            sh.cell(r, 1, k).font = BOLD
            sh.cell(r, 1).alignment = WRAP
            cc = sh.cell(r, 2, v)
            cc.alignment, cc.font = WRAP, Font(size=11)
            if fill:
                cc.fill = PatternFill("solid", fgColor=FILL[fill])
            sh.row_dimensions[r].height = max(est_h(v), 18)
            cur[0] += 1

        row(L["result"], status, status)
        if consolidated:
            if not rec:
                row(L["src"], L["never"], "NOT EXECUTED")
            elif rec.get("carried"):
                chg = rec.get("spec_changed")
                row(L["src"], L["carried"].format(run=rec.get("src_run", ""), when=rec.get("started", "")) + (L["spec_chg"] if chg else L["spec_unk"] if chg is None else ""),
                    "PENDING" if chg else None)
            else:
                row(L["src"], L["this_run"].format(run=rec.get("src_run", "")))
        row(L["when"], rec.get("started", "") + (f" ～ {rec.get('finished', '')[11:]}" if rec.get("finished") else ""))
        row(L["env"], env_text)
        row(L["url"], url)
        tms = rec.get("timings") or []
        if tms:
            row(L["dur"], " / ".join(f"{x['slot']}: {x['seconds']}s ({x['start']}–{x['end']})" for x in tms) + f" | Σ {round(sum(x['seconds'] for x in tms), 1)}s")
        cls, typ = c.get("cls", ""), c.get("typ", "")
        if tr is not None:
            cls, typ = getattr(tr, "CLS", {}).get(cls, cls), getattr(tr, "TYP", {}).get(typ, typ)
        row(L["clstyp"], f"{cls} / {typ}")
        row(L["role"], str(sp["role"]))
        row(L["pre"], str(sp["pre"]))
        row(L["acts"], "\n".join(map(str, sp["act"])))
        row(L["exp"], str(sp["exp"]))
        n_total = sum(len(step_images(s_, j)) for j, s_ in enumerate(steps, 1))
        row(L["total"], L["spec_cnt"].format(a=len(sp["act"]), b=n_total))
        vs = []
        for i, s in enumerate(steps, 1):
            cap = getattr(tr, "CAP", {}).get((no, i), s["caption"]) if tr is not None else s["caption"]
            act = getattr(tr, "ACT", {}).get((no, i), s["actual"]) if tr is not None else s["actual"]
            try:
                idx = c["act_lines"].index(s["action"])
            except ValueError:
                idx = min(int(s.get("orig_step") or 1), max(len(sp["act"]), 1)) - 1
            action_txt = sp["act"][idx] if sp["act"] else s["action"]
            vs.append(dict(s, cap=cap))
            row(f"{L['step']} {i}", f"{s['status']}：{cap}\n{L['op']}: {action_txt}\n{L['actual']}: {act}",
                s["status"] if s["status"] in ("FAIL", "BLOCKED", "PENDING") else None)
        rem = (getattr(tr, "REM", {}).get(no) if tr is not None else None) or rec.get("remark")
        if rem:
            row(L["note"], rem)
        cur[0] += 1
        n_img = 0
        for i, s in enumerate(vs, 1):
            ims = step_images(s, i)
            for j, im_ in enumerate(ims):
                if not os.path.exists(im_["annotated"]):
                    continue
                path, note = im_["annotated"], im_["note"]
                if tr is not None:
                    note = s["cap"] if j == len(ims) - 1 else getattr(tr, "NOTE", {}).get(im_["note"], im_["note"])
                    path = str(lang_dir / str(no) / os.path.basename(im_["annotated"]).replace("_annotated", f"_annotated_{lang}"))
                    _recaption(im_["annotated"], path, f"{L['step']} {im_['tag']}: {note}")
                r = cur[0]
                sh.cell(r, 1, f"{L['screen']} ({L['step']} {im_['tag']})").font = BOLD
                sh.cell(r, 2, os.path.basename(path) if cfg.show_filename else f"{L['step']} {im_['tag']}: {note}").font = Font(size=11)
                im = XlImage(path)
                w, h = Image.open(path).size
                im.width, im.height = 900, round(900 * h / w)
                sh.add_image(im, f"B{r + 1}")
                cur[0] += max(31, round(900 * h / w / 20) + 2)
                n_img += 1
        summary.append((no, f"No.{no}", sp["view"], status, rec.get("started", ""), L["n_img"].format(n=n_img),
                        (f"[{L['src']}: run {rec.get('src_run', '')}{' ※' if rec.get('spec_changed') else ''}] " if rec.get("carried") else "")
                        + " / ".join(f"{i}:{s['status']}：{s['cap']}" for i, s in enumerate(vs, 1))[:1000]))
    sm = wb.create_sheet(L["summary"], 0)
    for i, w in enumerate([12, 56, 14, 20, 40, 100], 1):
        sm.column_dimensions[chr(64 + i)].width = w
    cnt = {k: sum(1 for x in summary if x[3] == k) for k in FILL}
    sm["A1"] = L["sum_title"]
    sm["A1"].font = Font(bold=True, size=13)
    sm["A2"] = f"{L['sum_env']}：{env_text} ({url})"
    sm["A3"] = L["sum_note"]
    sm["A4"] = f"{L['agg']}：{len(summary)} — " + " / ".join(f"{k} {cnt[k]}" for k in FILL)
    if consolidated:
        recs = [RES.get(str(x[0]), {}) for x in summary]
        sm["A5"] = L["cons_note"].format(a=sum(1 for r in recs if r and not r.get("carried")), b=sum(1 for r in recs if r.get("carried")),
                                         c=sum(1 for r in recs if r.get("carried") and r.get("spec_changed")), n=sum(1 for r in recs if not r))
        sm["A5"].font = BOLD
    for i, hname in enumerate(L["cols"], 1):
        c_ = sm.cell(6, i, hname)
        c_.font, c_.fill = BOLD, PatternFill("solid", fgColor="FFFFF2CC")
    for j, row_ in enumerate(summary, 7):
        for i, v in enumerate(row_[1:], 1):
            c_ = sm.cell(j, i, v)
            c_.alignment = WRAP
            if i == 3:
                c_.fill = PatternFill("solid", fgColor=FILL.get(v, "FFEDEDED"))
        sm.row_dimensions[j].height = 64
    base = (os.path.splitext(os.path.basename(spec_path))[0] if spec_path else (cfg.suite or "suite")) + f"_evidence_{lang}" + ("_tong_hop" if consolidated else "") + ".xlsx"
    cfg.ensure()
    dst = str(cfg.excel_dir / base)
    try:
        wb.save(dst)
    except PermissionError:      # open in Excel: never kill it, save next to it
        dst = dst.replace(".xlsx", datetime.datetime.now().strftime("_%H%M%S.xlsx"))
        wb.save(dst)
        print("NOTE: target is open / locked, saved as a new file instead")
    print("saved", dst, cnt)
    return dst


if __name__ == "__main__":
    import argparse, importlib.util, sys
    from ..core.config import WORK
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Build the evidence workbook from data/results_<suite>.json")
    ap.add_argument("--suite", required=True)
    ap.add_argument("--spec", help="testcase workbook (omit for suites without a spec)")
    ap.add_argument("--lang", default=(cfg.langs or ["en"])[0], choices=sorted(LABELS))
    ap.add_argument("--tr", help="python file with translations for --lang")
    ap.add_argument("--only-run", action="store_true", help="only the cases that have a result in this run")
    ap.add_argument("--consolidated", action="store_true", help="results of tools.runs consolidate (QA_RESULTS_DIR) -> ..._tong_hop.xlsx")
    a = ap.parse_args()
    cfg.set_suite(a.suite)
    tr = None
    if a.tr:
        s_ = importlib.util.spec_from_file_location("tr_mod", a.tr)
        tr = importlib.util.module_from_spec(s_); s_.loader.exec_module(tr)
    spec = a.spec and (a.spec if os.path.isabs(a.spec) else str(WORK / a.spec))
    print(build(spec, lang=a.lang, tr=tr, only_run=a.only_run, consolidated=a.consolidated))
