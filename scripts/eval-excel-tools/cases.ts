import { Workbook, parseCell, colName } from "./workbook.js";

export interface Call { tool: string; args: any; ok: boolean; error?: string | null }
export interface Run { calls: Call[]; reply: string; before: Map<string, string>; wb: Workbook }
export interface Verdict { tool: boolean; state: boolean; reply: boolean; notes: string[] }
export interface Oracle { calls: Array<{ tool: string; args: Record<string, unknown> }>; reply: string }
export interface Case {
  id: string; tag: string; source: "real" | "synth"; prompt: string;
  fixture: () => Workbook; selection?: { sheet: string; address: string };
  /** Claims a judge must confirm in the reply (only for cases whose reply cannot be checked by code). */
  claims?: string[];
  check(run: Run): Verdict;
  oracle: Oracle;
}

/* ------------------------------ fixtures ------------------------------ */
const products = [["Táo", 3, 12000], ["Lê", 5, 15000], ["Chuối", 2, 8000], ["Nho", 8, 30000], ["Xoài", 4, 25000], ["Dưa hấu", 1, 40000], ["Ổi", 6, 10000], ["Mận", 7, 18000], ["Khế", 2, 9000], ["Bưởi", 9, 22000]] as const;
const revenue = [450000, 480000, 520000, 500000, 540000, 560000, 580000, 600000, 620000, 650000];
const cost = [300000, 310000, 320000, 330000, 340000, "#DIV/0!", 360000, 370000, 380000, 390000];
const fSales = (name = "Sales.xlsx") => { const w = new Workbook(name); w.addSheet("Data", [["Product", "Qty", "Price", "Total"], ...products.map(p => [...p, null])]); return w; };
const addRevenue = (w: Workbook) => w.addSheet("Doanh thu", [["Tháng", "Doanh thu", "Chi phí"], ...revenue.map((r, i) => [`Tháng ${i + 1}`, r, typeof cost[i] === "string" ? { err: cost[i] as string } : cost[i]!])]);
const fReport = () => { const w = new Workbook("Report.xlsx"); addRevenue(w); return w; };
const fNotes = () => { const w = new Workbook("Notes.xlsx"); w.addSheet("Ghi chú", [["Họp nhóm"], ["Thág 3 báo cáo"], ["Hà Nôi"], ["Gửi email"], ["Ghi chú cuối"], ["Kế hoạch"]]); return w; };
const fOrders = () => { const w = new Workbook("Orders.xlsx"); w.addSheet("Orders", [["Order", "Product", "Qty", "Price"], ["DH001", "Táo", 3, 12000], ["DH002", "Lê", 5, 15000], ["DH003", "Nho", 8, 30000], ["DH004", "Xoài", -4, 25000], ["DH005", "Ổi", 6, 10000], ["DH006", "Mận", 7, "N/A"], ["DH007", "Khế", 2, 9000], ["DH008", "Bưởi", 9, 22000], ["DH009", "Chuối", 2, 8000], ["DH003", "Nho", 8, 30000], ["DH011", "Dưa hấu", 1, 40000], ["DH012", "Cam", 4, 15000]]); return w; };

/* ------------------------------- helpers ------------------------------ */
const WRITES = new Set(["write_range", "set_formula", "clear_range", "insert_rows", "modify_structure", "manage_sheet", "modify_chart", "find_replace", "set_cells", "copy_to", "sort_range", "merge_range", "execute_office_js"]);
const wroteOk = (r: Run) => r.calls.some(c => WRITES.has(c.tool) && c.ok);
const called = (r: Run, ...tools: string[]) => r.calls.some(c => tools.includes(c.tool));
const unchanged = (r: Run) => same(r.before, r.wb.snapshot());
const same = (a: Map<string, string>, b: Map<string, string>) => a.size === b.size && [...a].every(([k, v]) => b.get(k) === v);
/** Keys whose raw content differs between two snapshots (ignores "#"-prefixed meta keys). */
const diff = (a: Map<string, string>, b: Map<string, string>) => [...new Set([...a.keys(), ...b.keys()])].filter(k => a.get(k) !== b.get(k));
const changed = (r: Run) => diff(r.before, r.wb.snapshot());
const only = (r: Run, allowed: (k: string) => boolean) => changed(r).every(k => allowed(k));
const raw = (r: Run, sheet: string, cell: string) => { const s = r.wb.sheets.find(x => x.name === sheet); const c = s?.cells.get(cell); return c ? (c.f !== undefined ? c.f : c.v) : null; };
const val = (r: Run, sheet: string, cell: string) => { const s = r.wb.sheets.find(x => x.name === sheet); return s ? r.wb.value(s, cell) : null; };
const isFormula = (r: Run, sheet: string, cell: string) => r.wb.sheets.find(x => x.name === sheet)?.cells.get(cell)?.f !== undefined;
const norm = (s: unknown) => String(s ?? "").normalize("NFC").trim().toLowerCase();

/** Numbers in a reply, with thousands separators and "triệu / nghìn / k / million" read as a person would. */
export function numbersIn(text: string): number[] {
  const out: number[] = [];
  for (const m of text.matchAll(/(\d[\d.,]*)(\s*(triệu|million|nghìn|ngàn|k\b))?/gi)) {
    let raw = m[1]!.replace(/[.,]+$/, ""); const unit = (m[3] ?? "").toLowerCase();
    const n = /^\d{1,3}([.,]\d{3})+$/.test(raw) ? Number(raw.replace(/[.,]/g, "")) : Number(raw.replace(",", "."));
    if (Number.isNaN(n)) continue; out.push(unit === "triệu" || unit === "million" ? n * 1e6 : unit ? n * 1e3 : n);
  }
  return out;
}
const hasNumber = (text: string, n: number) => numbersIn(text).some(x => Math.abs(x - n) < 0.5);
/** Rows of sheet `name` at/after `first` moved down by `count`, rows before `first` untouched, inserted rows empty, nothing else changed. */
function shifted(r: Run, name: string, first: number, count: number): boolean {
  const after = r.wb.snapshot(); const expected = new Map<string, string>();
  for (const [k, v] of r.before) {
    const m = k.startsWith(`${name}!`) && /^(.+)!([A-Z]+)(\d+)$/.exec(k);
    if (!m) { expected.set(k, v); continue; }
    const row = Number(m[3]); expected.set(row >= first ? `${name}!${m[2]}${row + count}` : k, v);
  }
  return same(expected, after);
}
const reply = (r: Run) => r.reply;

const V = (tool: boolean, state: boolean, rep: boolean, ...notes: string[]): Verdict => ({ tool, state, reply: rep, notes });

/* -------------------------------- cases -------------------------------- */
export const CASES: Case[] = [
  { id: "R1", tag: "read", source: "real", prompt: "Read all the data in this workbook. For each worksheet, list the sheet name, the used range address, and the contents of the used range (values, not formulas).",
    fixture: () => { const w = fSales(); addRevenue(w); return w; },
    check: r => V(called(r, "get_sheets", "read_range", "get_cell_ranges") && !wroteOk(r), unchanged(r), /Data/.test(reply(r)) && /Doanh thu/.test(reply(r)) && /A1:D11/.test(reply(r)) && /A1:C11/.test(reply(r))),
    oracle: { calls: [{ tool: "get_sheets", args: {} }, { tool: "read_range", args: { address: "A1:D11" } }], reply: "Sheet Data: A1:D11 ... Sheet Doanh thu: A1:C11 ..." } },
  { id: "R2", tag: "read", source: "real", prompt: "đang chọn gì", fixture: fSales, selection: { sheet: "Data", address: "B3:C4" },
    check: r => V(!wroteOk(r), unchanged(r), /B3/.test(reply(r)) && /C4/.test(reply(r))),
    oracle: { calls: [], reply: "Bạn đang chọn B3:C4 trên sheet Data." } },
  { id: "R3", tag: "read", source: "synth", prompt: "Có ô nào đang bị lỗi công thức không?", fixture: fReport,
    check: r => V(called(r, "find_errors", "search_data", "read_range", "get_cell_ranges") && !wroteOk(r), unchanged(r), /C7/.test(reply(r))),
    oracle: { calls: [{ tool: "find_errors", args: {} }], reply: "Ô C7 trên sheet Doanh thu đang lỗi #DIV/0!." } },
  { id: "R4", tag: "read", source: "real", prompt: "bạn làm được gì", fixture: fSales,
    check: r => V(r.calls.length === 0, unchanged(r), reply(r).trim().length > 20),
    oracle: { calls: [], reply: "Tôi có thể đọc, sửa, tính toán và lập báo cáo trong workbook của bạn." } },
  { id: "P1", tag: "python", source: "real", prompt: "Use the run_python tool to compute 6*7 and tell me the result.", fixture: fSales,
    check: r => V(r.calls.some(c => c.tool === "run_python" && c.ok), unchanged(r), /\b42\b/.test(reply(r))),
    oracle: { calls: [{ tool: "run_python", args: { code: "print(6*7)" } }], reply: "Kết quả là 42." } },
  { id: "W1", tag: "write", source: "real", prompt: "sửa lỗi chính ta đi", fixture: fNotes, selection: { sheet: "Ghi chú", address: "A1:A6" },
    check: r => V(wroteOk(r), norm(raw(r, "Ghi chú", "A2")) === "tháng 3 báo cáo" && norm(raw(r, "Ghi chú", "A3")) === "hà nội" && only(r, k => k === "Ghi chú!A2" || k === "Ghi chú!A3"), reply(r).trim().length > 0),
    oracle: { calls: [{ tool: "write_range", args: { address: "A2:A3", values: [["Tháng 3 báo cáo"], ["Hà Nội"]], allowOverwrite: true } }], reply: "Đã sửa A2 và A3." } },
  { id: "W2", tag: "write", source: "synth", prompt: "Tính tổng cột B vào ô B12", fixture: fReport,
    check: r => V(wroteOk(r), isFormula(r, "Doanh thu", "B12") && /SUM\(/i.test(String(raw(r, "Doanh thu", "B12"))) && val(r, "Doanh thu", "B12") === 5500000 && only(r, k => k === "Doanh thu!B12"), true),
    oracle: { calls: [{ tool: "set_formula", args: { address: "B12", formula: "=SUM(B2:B11)" } }], reply: "Đã đặt B12 = SUM(B2:B11) = 5.500.000." } },
  { id: "W3", tag: "write", source: "synth", prompt: "Điền cột D = Qty × Price cho tất cả các dòng", fixture: fSales,
    check: r => V(wroteOk(r), products.every((p, i) => isFormula(r, "Data", `D${i + 2}`) && val(r, "Data", `D${i + 2}`) === p[1] * p[2]) && norm(raw(r, "Data", "D1")) === "total" && only(r, k => /^Data!D([2-9]|1[01])$/.test(k)), true),
    oracle: { calls: [{ tool: "set_formula", args: { address: "D2:D11", formula: "=B2*C2" } }], reply: "Đã điền D2:D11 = Qty × Price." } },
  { id: "W4", tag: "write", source: "synth", prompt: 'Ghi chữ "n/a" vào C2:C4', fixture: fSales,
    claims: ["The reply says C2:C4 already contain data (prices).", "The reply asks the user to confirm before overwriting, or says it did not overwrite."],
    check: r => V(true, unchanged(r), true),
    oracle: { calls: [], reply: "C2:C4 đã có giá trị (12000, 15000, 8000). Bạn có chắc muốn ghi đè bằng n/a không?" } },
  { id: "W5", tag: "write", source: "synth", prompt: "Cột C nhìn rối quá", fixture: fSales,
    check: r => { const s = r.wb.sheets[0]!; const vals = [...r.before].filter(([k]) => !k.includes("#")).every(([k, v]) => r.wb.snapshot().get(k) === v); return V(!r.calls.some(c => (c.tool === "modify_structure" && ["hide", "delete"].includes(c.args?.operation)) || c.tool === "clear_range"), vals && s.hiddenCols.size === 0 && s.hiddenRows.size === 0, reply(r).trim().length > 0); },
    oracle: { calls: [{ tool: "format_range", args: { address: "C1:C11", format: { horizontalAlignment: "Right" } } }], reply: "Tôi đã căn phải cột C cho dễ nhìn." } },
  { id: "I1", tag: "insert", source: "synth", prompt: "Chèn 3 dòng trống phía trên dòng 5", fixture: fSales,
    check: r => V(called(r, "insert_rows", "modify_structure"), shifted(r, "Data", 5, 3), true),
    oracle: { calls: [{ tool: "insert_rows", args: { at: "5", count: 3, position: "above" } }], reply: "Đã chèn 3 dòng trống trên dòng 5." } },
  { id: "I2", tag: "insert", source: "synth", prompt: "Insert one blank row right below the header row.", fixture: fSales,
    check: r => V(called(r, "insert_rows", "modify_structure"), shifted(r, "Data", 2, 1), true),
    oracle: { calls: [{ tool: "insert_rows", args: { at: "1", count: 1, position: "below" } }], reply: "Inserted one blank row under the header." } },
  { id: "I3", tag: "insert", source: "synth", prompt: "Add 2 blank rows below A10.", fixture: fSales,
    check: r => V(called(r, "insert_rows", "modify_structure"), shifted(r, "Data", 11, 2), true),
    oracle: { calls: [{ tool: "insert_rows", args: { at: "A10", count: 2, position: "below" } }], reply: "Added 2 blank rows below row 10." } },
  { id: "I4", tag: "insert", source: "synth", prompt: "Thêm 1 dòng trống trên dòng 3 của sheet Doanh thu", fixture: () => { const w = fSales(); addRevenue(w); return w; },
    check: r => V(r.calls.some(c => c.tool === "insert_rows" && c.ok), shifted(r, "Doanh thu", 3, 1), true),
    oracle: { calls: [{ tool: "insert_rows", args: { sheet: { name: "Doanh thu" }, at: "3", count: 1, position: "above" } }], reply: "Đã thêm 1 dòng trống trên dòng 3 của Doanh thu." } },
  { id: "I5", tag: "insert", source: "synth", prompt: "Thêm 1500 dòng trống dưới dòng 2", fixture: fSales,
    claims: ["The reply states the 1000-row limit or the total number of rows inserted.", "The reply does not claim more rows were inserted than were."],
    check: r => { const ok = r.calls.filter(c => c.ok && (c.tool === "insert_rows" || (c.tool === "modify_structure" && c.args?.operation === "insert" && c.args?.dimension === "rows"))); const n = ok.reduce((t, c) => t + (c.args?.count ?? 1), 0); return V(true, [0, 1000, 1500].includes(n) && shifted(r, "Data", 3, n), true); },
    oracle: { calls: [{ tool: "insert_rows", args: { at: "2", count: 1000, position: "below" } }, { tool: "insert_rows", args: { at: "2", count: 500, position: "below" } }], reply: "Đã chèn tổng cộng 1500 dòng (giới hạn mỗi lần là 1000, nên tôi chia làm hai lần)." } },
  { id: "I6", tag: "insert", source: "synth", prompt: "Chèn thêm vài dòng", fixture: fSales,
    claims: ["The reply asks which row to insert at, or how many rows.", "The reply does not say any rows were inserted."],
    check: r => V(!called(r, "insert_rows", "modify_structure"), unchanged(r), true),
    oracle: { calls: [], reply: "Bạn muốn chèn ở dòng nào và chèn bao nhiêu dòng?" } },
  { id: "I7", tag: "insert", source: "synth", prompt: "Xóa dòng 4", fixture: fSales,
    check: r => { const del = (() => { const e = new Map<string, string>(); for (const [k, v] of r.before) { const m = /^Data!([A-Z]+)(\d+)$/.exec(k); if (!m) { e.set(k, v); continue; } const row = Number(m[2]); if (row === 4) continue; e.set(row > 4 ? `Data!${m[1]}${row - 1}` : k, v); } return e; })(); const s = r.wb.snapshot(); return V(!called(r, "insert_rows"), unchanged(r) || same(del, s), reply(r).trim().length > 0); },
    oracle: { calls: [{ tool: "modify_structure", args: { operation: "delete", dimension: "rows", reference: "4" } }], reply: "Đã xóa dòng 4." } },
  { id: "A1", tag: "add-data", source: "synth", prompt: "Thêm sản phẩm mới: Cam, số lượng 5, giá 15000 vào cuối bảng", fixture: fSales,
    check: r => V(wroteOk(r), raw(r, "Data", "A12") === "Cam" && raw(r, "Data", "B12") === 5 && raw(r, "Data", "C12") === 15000 && only(r, k => /^Data![A-D]12$/.test(k)), true),
    oracle: { calls: [{ tool: "write_range", args: { address: "A12:C12", values: [["Cam", 5, 15000]] } }], reply: "Đã thêm Cam ở dòng 12." } },
  { id: "A2", tag: "add-data", source: "synth", prompt: "Add three more rows of sample data in the same format as the existing ones.", fixture: fSales,
    check: r => V(wroteOk(r), [12, 13, 14].every(n => typeof raw(r, "Data", `A${n}`) === "string" && typeof raw(r, "Data", `B${n}`) === "number" && typeof raw(r, "Data", `C${n}`) === "number") && only(r, k => /^Data![A-C](12|13|14)$/.test(k)), true),
    oracle: { calls: [{ tool: "write_range", args: { address: "A12:C14", values: [["Dừa", 3, 11000], ["Mít", 2, 20000], ["Ổi đỏ", 4, 9500]] } }], reply: "Added three rows (12-14)." } },
  { id: "A3", tag: "add-data", source: "synth", prompt: 'Thêm cột Ghi chú ở cuối bảng, ghi "Còn hàng" cho mọi dòng', fixture: fSales,
    check: r => V(wroteOk(r), norm(raw(r, "Data", "E1")) === "ghi chú" && Array.from({ length: 10 }, (_, i) => norm(raw(r, "Data", `E${i + 2}`)) === "còn hàng").every(Boolean) && only(r, k => /^Data!E([1-9]|1[01])$/.test(k)), true),
    oracle: { calls: [{ tool: "write_range", args: { address: "E1:E11", values: [["Ghi chú"], ...Array.from({ length: 10 }, () => ["Còn hàng"])] } }], reply: "Đã thêm cột E." } },
  { id: "A4", tag: "add-data", source: "synth", prompt: "Nhập doanh thu tháng 11 là 52 triệu", fixture: fReport,
    check: r => V(wroteOk(r), norm(raw(r, "Doanh thu", "A12")).includes("11") && [52000000, 52000].includes(raw(r, "Doanh thu", "B12") as number) && only(r, k => /^Doanh thu!.12$/.test(k)), true),
    oracle: { calls: [{ tool: "write_range", args: { address: "A12:B12", values: [["Tháng 11", 52000000]] } }], reply: "Đã nhập 52.000.000 cho Tháng 11." } },
  { id: "N1", tag: "analysis", source: "synth", prompt: "Tổng và trung bình doanh thu cả năm là bao nhiêu?", fixture: fReport,
    check: r => V(!wroteOk(r), unchanged(r), hasNumber(reply(r), 5500000) && hasNumber(reply(r), 550000)),
    oracle: { calls: [{ tool: "read_range", args: { address: "B2:B11" } }], reply: "Tổng doanh thu là 5.500.000 và trung bình là 550.000 mỗi tháng." } },
  { id: "N2", tag: "analysis", source: "synth", prompt: "Tháng nào doanh thu cao nhất, tháng nào thấp nhất?", fixture: fReport,
    check: r => V(!wroteOk(r), unchanged(r), /(tháng|month|t)\s*10(?!\d)/i.test(reply(r)) && /(tháng|month|t)\s*1(?!\d)/i.test(reply(r))),
    oracle: { calls: [{ tool: "read_range", args: { address: "A2:B11" } }], reply: "Cao nhất là Tháng 10 (650.000), thấp nhất là Tháng 1 (450.000)." } },
  { id: "N3", tag: "analysis", source: "synth", prompt: "Thêm cột Lợi nhuận = Doanh thu - Chi phí", fixture: fReport,
    claims: ["The reply says the profit in row 7 (D7) is an error.", "The reply says the error comes from the cost cell C7."],
    check: r => V(wroteOk(r), Array.from({ length: 10 }, (_, i) => { const row = i + 2; return isFormula(r, "Doanh thu", `D${row}`) && (row === 7 ? /^#/.test(String(val(r, "Doanh thu", "D7"))) : val(r, "Doanh thu", `D${row}`) === revenue[i]! - (cost[i] as number)); }).every(Boolean) && only(r, k => /^Doanh thu!D([1-9]|1[01])$/.test(k)), true),
    oracle: { calls: [{ tool: "write_range", args: { address: "D1:D11", values: [["Lợi nhuận"], ...Array.from({ length: 10 }, (_, i) => [`=B${i + 2}-C${i + 2}`])] } }], reply: "Đã thêm cột D. Lưu ý D7 báo lỗi #DIV/0! vì chi phí ở C7 đang lỗi." } },
  { id: "N4", tag: "analysis", source: "synth", prompt: "Đơn hàng có giá trị lớn nhất là bao nhiêu (số lượng × giá)?", fixture: fSales,
    check: r => V(!wroteOk(r), unchanged(r), hasNumber(reply(r), 240000)),
    oracle: { calls: [{ tool: "read_range", args: { address: "A2:C11" } }], reply: "Lớn nhất là Nho: 8 × 30.000 = 240.000." } },
  { id: "N5", tag: "analysis", source: "synth", prompt: "Kiểm tra dữ liệu giúp tôi, có gì bất thường không?", fixture: fOrders,
    claims: ["The reply identifies a negative quantity as a problem (cell C5).", "The reply identifies the text value in a price cell as a problem (D7).", "The reply identifies the repeated order id DH003 as a problem."],
    check: r => V(!wroteOk(r), unchanged(r), true),
    oracle: { calls: [{ tool: "read_range", args: { address: "A1:D13" } }], reply: "Có 3 vấn đề: số lượng âm ở C5; giá là chữ N/A ở D7; mã đơn DH003 bị trùng (A4 và A11)." } },
  { id: "B1", tag: "report", source: "synth", prompt: "Lập báo cáo tóm tắt doanh thu: tổng, trung bình, tháng cao nhất và thấp nhất.", fixture: fReport,
    check: r => { const vals = r.wb.sheets.flatMap(s => [...s.cells.keys()].filter(k => s.cells.get(k)!.f !== undefined).map(k => r.wb.value(s, k))); const after = r.wb.snapshot(); const src = [...r.before].filter(([k]) => !k.includes("#")).every(([k, v]) => after.get(k) === v);
      return V(wroteOk(r), [5500000, 550000, 650000, 450000].every(n => vals.includes(n)) && src, reply(r).trim().length > 0); },
    oracle: { calls: [{ tool: "manage_sheet", args: { operation: "add", name: "Báo cáo" } }, { tool: "write_range", args: { sheet: { name: "Báo cáo" }, address: "A1:B4", values: [["Tổng", "=SUM('Doanh thu'!B2:B11)"], ["Trung bình", "=AVERAGE('Doanh thu'!B2:B11)"], ["Cao nhất", "=MAX('Doanh thu'!B2:B11)"], ["Thấp nhất", "=MIN('Doanh thu'!B2:B11)"]] } }], reply: "Đã tạo sheet Báo cáo với các công thức; đã kiểm tra giá trị." } },
  { id: "B2", tag: "report", source: "synth", prompt: "Tạo báo cáo doanh thu theo quý ở một sheet mới", fixture: fReport,
    claims: ["The reply says the last quarter (Q4) covers only one month."],
    check: r => { const s = r.wb.sheets.find(x => x.name !== "Doanh thu"); const vals = s ? [...s.cells.keys()].map(k => r.wb.value(s, k)) : []; const f = s ? [...s.cells.values()].filter(c => c.f !== undefined).length : 0; const src = [...r.before].filter(([k]) => k.startsWith("Doanh thu!")).every(([k, v]) => r.wb.snapshot().get(k) === v);
      return V(r.calls.some(c => c.tool === "manage_sheet" && c.ok), !!s && f >= 3 && [1450000, 1600000, 1800000, 650000].every(n => vals.includes(n)) && src, true); },
    oracle: { calls: [{ tool: "manage_sheet", args: { operation: "add", name: "Theo quý" } }, { tool: "write_range", args: { sheet: { name: "Theo quý" }, address: "A1:B4", values: [["Q1", "=SUM('Doanh thu'!B2:B4)"], ["Q2", "=SUM('Doanh thu'!B5:B7)"], ["Q3", "=SUM('Doanh thu'!B8:B10)"], ["Q4", "=SUM('Doanh thu'!B11:B11)"]] } }], reply: "Đã tạo báo cáo theo quý. Lưu ý quý 4 mới có một tháng (tháng 10)." } },
  { id: "B3", tag: "report", source: "synth", prompt: "Vẽ biểu đồ cột doanh thu theo tháng", fixture: fReport,
    check: r => V(r.calls.some(c => c.tool === "modify_chart" && c.ok && c.args?.operation === "create" && /col|bar/i.test(String(c.args?.chartType)) && /A1:B11/i.test(String(c.args?.source))), unchanged(r), reply(r).trim().length > 0),
    oracle: { calls: [{ tool: "modify_chart", args: { operation: "create", sheet: { name: "Doanh thu" }, chartType: "columnClustered", source: "'Doanh thu'!A1:B11", title: "Doanh thu theo tháng" } }], reply: "Đã vẽ biểu đồ cột." } },
];

/** The selection snapshot a pane would send for a case (the whole first cell when nothing is selected). */
export function selectionFor(c: Case, wb: Workbook) {
  const sel = c.selection ?? { sheet: wb.sheets[wb.active]!.name, address: "A1" };
  wb.selection = sel;
  const s = wb.sheet(sel.sheet); const [a, b] = sel.address.split(":"); const p = parseCell(a!), q = parseCell(b ?? a!);
  const values: unknown[][] = []; for (let r = p.row; r <= q.row; r++) { const row: unknown[] = []; for (let col = p.col; col <= q.col; col++) row.push(wb.value(s, `${colName(col)}${r}`)); values.push(row); }
  return { host: "excel" as const, workbookName: wb.name, sheet: s.name, sheetId: null, address: `${s.name}!${sel.address}`, rowCount: q.row - p.row + 1, columnCount: q.col - p.col + 1, values, headers: null, truncated: false };
}
