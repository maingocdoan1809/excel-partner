/**
 * A small in-memory workbook that stands in for Office.js in the Excel-tools eval: enough of the real tools'
 * behaviour (formulas, would_overwrite, insert_rows, errors, structure) to grade the assistant on tool calls
 * and on the workbook it leaves behind. It is not a spreadsheet engine; unknown tools answer "unsupported".
 */

export type Cell = { v: string | number | boolean | null; f?: string };
export interface Sheet { name: string; cells: Map<string, Cell>; hiddenRows: Set<number>; hiddenCols: Set<number> }
export interface ToolResultLike { ok: boolean; data?: unknown; error?: { code: string; message: string; detail?: unknown } }

const ERR = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!)$/;
export const colNum = (c: string) => [...c.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const colName = (n: number) => { let s = ""; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const key = (col: number, row: number) => `${colName(col)}${row}`;
export function parseCell(a: string): { col: number; row: number } { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim()); if (!m) throw new Error(`bad cell ${a}`); return { col: colNum(m[1]!), row: Number(m[2]) }; }
export function parseRange(a: string): { c1: number; r1: number; c2: number; r2: number } {
  const [x, y] = a.replace(/\$/g, "").split(":"); const p = parseCell(x!), q = parseCell(y ?? x!);
  return { c1: Math.min(p.col, q.col), r1: Math.min(p.row, q.row), c2: Math.max(p.col, q.col), r2: Math.max(p.row, q.row) };
}

export class Workbook {
  sheets: Sheet[] = [];
  name = "Book.xlsx";
  active = 0;
  charts: Array<Record<string, unknown>> = [];
  unsimulated: string[] = [];
  selection: { sheet: string; address: string } | null = null;

  constructor(name: string) { this.name = name; }
  addSheet(name: string, rows: Array<Array<Cell["v"] | { f: string } | { err: string }>> = []): Sheet {
    const sheet: Sheet = { name, cells: new Map(), hiddenRows: new Set(), hiddenCols: new Set() };
    rows.forEach((row, r) => row.forEach((v, c) => {
      if (v === null || v === undefined) return;
      if (typeof v === "object" && "f" in v) sheet.cells.set(key(c + 1, r + 1), { v: null, f: v.f });
      else if (typeof v === "object" && "err" in v) sheet.cells.set(key(c + 1, r + 1), { v: v.err });
      else sheet.cells.set(key(c + 1, r + 1), { v });
    }));
    this.sheets.push(sheet); return sheet;
  }
  clone(): Workbook {
    const w = new Workbook(this.name); w.active = this.active; w.selection = this.selection && { ...this.selection };
    w.sheets = this.sheets.map(s => ({ name: s.name, cells: new Map([...s.cells].map(([k, c]) => [k, { ...c }])), hiddenRows: new Set(s.hiddenRows), hiddenCols: new Set(s.hiddenCols) }));
    return w;
  }
  sheet(ref?: { name?: string; id?: string } | string | null): Sheet {
    const name = typeof ref === "string" ? ref : ref?.name;
    const s = name ? this.sheets.find(x => x.name.toLowerCase() === name.toLowerCase()) : this.sheets[this.active];
    if (!s) throw new ToolError("not_found", `No sheet named '${name}'. Sheets: ${this.sheets.map(x => x.name).join(", ")}`);
    return s;
  }
  used(s: Sheet): { c2: number; r2: number } | null {
    let c2 = 0, r2 = 0; for (const k of s.cells.keys()) { const p = parseCell(k); c2 = Math.max(c2, p.col); r2 = Math.max(r2, p.row); }
    return c2 ? { c2, r2 } : null;
  }
  /** The value a user sees: the stored value, or the formula's result. */
  value(s: Sheet, k: string, depth = 0): string | number | boolean | null {
    const c = s.cells.get(k); if (!c) return null;
    if (c.f === undefined) return c.v;
    if (depth > 40) return "#REF!";
    return new Evaluator(this, s, depth + 1).run(c.f);
  }
  isEmpty(s: Sheet, k: string) { const c = s.cells.get(k); return !c || (c.f === undefined && (c.v === null || c.v === "")); }
  /** Every non-empty cell as "Sheet!A1" -> raw (formula text or value), used to diff before and after. */
  snapshot(): Map<string, string> {
    const out = new Map<string, string>();
    for (const s of this.sheets) {
      out.set(`${s.name}!#hidden`, JSON.stringify([[...s.hiddenRows].sort(), [...s.hiddenCols].sort()]));
      for (const [k, c] of s.cells) if (!(c.f === undefined && (c.v === null || c.v === ""))) out.set(`${s.name}!${k}`, c.f !== undefined ? c.f : JSON.stringify(c.v));
    }
    out.set("#sheets", this.sheets.map(s => s.name).join("|"));
    return out;
  }
}

export class ToolError extends Error { constructor(public code: string, message: string, public detail?: unknown) { super(message); } }

/* ------------------------------ formulas ------------------------------ */

type V = string | number | boolean | null | V[][];
class Evaluator {
  private i = 0; private t: string[] = [];
  constructor(private wb: Workbook, private sheet: Sheet, private depth: number) {}
  run(formula: string): string | number | boolean | null {
    try {
      this.t = tokenize(formula.replace(/^=/, "")); this.i = 0;
      const v = this.cmp(); if (this.i < this.t.length) throw new Err("#NAME?");
      return this.scalar(v) as string | number | boolean | null;
    } catch (e) { return e instanceof Err ? e.code : "#VALUE!"; }
  }
  private peek() { return this.t[this.i]; }
  private next() { return this.t[this.i++]; }
  private scalar(v: V): V { return Array.isArray(v) ? (v[0]?.[0] ?? null) : v; }
  private num(v: V): number { const s = this.scalar(v); if (typeof s === "string" && ERR.test(s)) throw new Err(s); if (s === null || s === "") return 0; if (typeof s === "boolean") return s ? 1 : 0; const n = Number(s); if (Number.isNaN(n)) throw new Err("#VALUE!"); return n; }
  private cmp(): V {
    let l = this.cat();
    while (["=", "<>", "<", ">", "<=", ">="].includes(this.peek() ?? "")) {
      const op = this.next()!; const r = this.cat(); const a = this.scalar(l) as never, b = this.scalar(r) as never;
      l = op === "=" ? a === b : op === "<>" ? a !== b : op === "<" ? a < b : op === ">" ? a > b : op === "<=" ? a <= b : a >= b;
    }
    return l;
  }
  private cat(): V { let l = this.add(); while (this.peek() === "&") { this.next(); l = String(this.scalar(l) ?? "") + String(this.scalar(this.add()) ?? ""); } return l; }
  private add(): V { let l = this.mul(); while (this.peek() === "+" || this.peek() === "-") { const op = this.next(); const r = this.mul(); l = op === "+" ? this.num(l) + this.num(r) : this.num(l) - this.num(r); } return l; }
  private mul(): V { let l = this.pow(); while (this.peek() === "*" || this.peek() === "/") { const op = this.next(); const r = this.pow(); if (op === "*") l = this.num(l) * this.num(r); else { const d = this.num(r); if (d === 0) throw new Err("#DIV/0!"); l = this.num(l) / d; } } return l; }
  private pow(): V { let l = this.unary(); while (this.peek() === "^") { this.next(); l = this.num(l) ** this.num(this.unary()); } return l; }
  private unary(): V { if (this.peek() === "-") { this.next(); return -this.num(this.unary()); } if (this.peek() === "+") { this.next(); return this.unary(); } return this.atom(); }
  private atom(): V {
    const t = this.next(); if (t === undefined) throw new Err("#VALUE!");
    if (t === "(") { const v = this.cmp(); if (this.next() !== ")") throw new Err("#VALUE!"); return v; }
    if (/^\d/.test(t) || /^\.\d/.test(t)) return Number(t);
    if (t.startsWith('"')) return t.slice(1, -1).replace(/""/g, '"');
    if (/^(TRUE|FALSE)$/i.test(t)) return t.toUpperCase() === "TRUE";
    if (this.peek() === "(") { this.next(); const args: V[] = []; if (this.peek() !== ")") { for (;;) { args.push(this.cmp()); if (this.peek() === ",") { this.next(); continue; } break; } } if (this.next() !== ")") throw new Err("#VALUE!"); return this.fn(t.toUpperCase(), args); }
    return this.ref(t);
  }
  private ref(t: string): V {
    let sheet = this.sheet, a = t; const bang = t.lastIndexOf("!");
    if (bang >= 0) { const nm = t.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'"); const s = this.wb.sheets.find(x => x.name.toLowerCase() === nm.toLowerCase()); if (!s) throw new Err("#REF!"); sheet = s; a = t.slice(bang + 1); }
    if (!/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(a)) throw new Err("#NAME?");
    const { c1, r1, c2, r2 } = parseRange(a); const grid: V[][] = [];
    for (let r = r1; r <= r2; r++) { const row: V[] = []; for (let c = c1; c <= c2; c++) row.push(this.wb.value(sheet, key(c, r), this.depth) as V); grid.push(row); }
    return grid.length === 1 && grid[0]!.length === 1 ? (grid[0]![0] as V) : grid;
  }
  private flat(args: V[]): Array<string | number | boolean | null> { return args.flatMap(a => Array.isArray(a) ? a.flat() as never[] : [a]); }
  private nums(args: V[]): number[] { const out: number[] = []; for (const a of args) { const cells = Array.isArray(a) ? (a.flat() as never[]) : [a]; for (const c of cells) { if (typeof c === "string" && ERR.test(c)) throw new Err(c); if (typeof c === "number") out.push(c); else if (!Array.isArray(a) && c !== null && c !== "") out.push(this.num(c)); } } return out; }
  private fn(name: string, a: V[]): V {
    switch (name) {
      case "SUM": return this.nums(a).reduce((x, y) => x + y, 0);
      case "AVERAGE": { const n = this.nums(a); if (!n.length) throw new Err("#DIV/0!"); return n.reduce((x, y) => x + y, 0) / n.length; }
      case "MAX": { const n = this.nums(a); return n.length ? Math.max(...n) : 0; }
      case "MIN": { const n = this.nums(a); return n.length ? Math.min(...n) : 0; }
      case "COUNT": return this.flat(a).filter(x => typeof x === "number").length;
      case "COUNTA": return this.flat(a).filter(x => x !== null && x !== "").length;
      case "ABS": return Math.abs(this.num(a[0]!));
      case "ROUND": { const f = 10 ** this.num(a[1] ?? 0); return Math.round(this.num(a[0]!) * f) / f; }
      case "IF": return this.scalar(a[0]!) ? (a[1] ?? true) : (a[2] ?? false);
      case "IFERROR": { const v = this.scalar(a[0]!); return typeof v === "string" && ERR.test(v) ? (a[1] ?? "") : v; }
      case "SUMPRODUCT": { const g = a.map(x => (Array.isArray(x) ? (x.flat() as number[]) : [x as number])); return g[0]!.reduce((s, _v, i) => s + g.reduce((p, col) => p * Number(col[i] ?? 0), 1), 0); }
      case "INDEX": { const g = a[0] as V[][]; const r = this.num(a[1] ?? 1), c = this.num(a[2] ?? 1) || 1; const grid = Array.isArray(g) ? g : [[g]]; const cell = grid.length === 1 && grid[0]!.length > 1 && !a[2] ? grid[0]![r - 1] : grid[r - 1]?.[c - 1]; if (cell === undefined) throw new Err("#REF!"); return cell as V; }
      case "MATCH": { const needle = this.scalar(a[0]!); const arr = this.flat([a[1]!]); const i = arr.findIndex(x => x === needle); if (i < 0) throw new Err("#N/A"); return i + 1; }
      case "SUMIF": { const r = this.flat([a[0]!]); const crit = this.scalar(a[1]!); const s = a[2] ? this.flat([a[2]]) : r; return r.reduce<number>((t, x, i) => t + (x === crit ? Number(s[i] ?? 0) : 0), 0); }
      case "AVERAGEIF": { const r = this.flat([a[0]!]); const crit = this.scalar(a[1]!); const s = a[2] ? this.flat([a[2]]) : r; const hit = r.map((x, i) => (x === crit ? Number(s[i] ?? 0) : null)).filter((x): x is number => x !== null); if (!hit.length) throw new Err("#DIV/0!"); return hit.reduce((x, y) => x + y, 0) / hit.length; }
      default: throw new Err("#NAME?");
    }
  }
}
class Err extends Error { constructor(public code: string) { super(code); } }
function tokenize(s: string): string[] {
  const out: string[] = []; let i = 0;
  while (i < s.length) {
    const ch = s[i]!;
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') { let j = i + 1; for (; j < s.length; j++) { if (s[j] === '"') { if (s[j + 1] === '"') j++; else break; } } out.push(s.slice(i, j + 1)); i = j + 1; continue; }
    const m = /^(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])|^[A-Za-z_][\w.]*|^\d+\.?\d*(?:[eE][+-]?\d+)?|^\.\d+/.exec(s.slice(i));
    if (m) { out.push(m[0]); i += m[0].length; continue; }
    const two = s.slice(i, i + 2); if (["<>", "<=", ">="].includes(two)) { out.push(two); i += 2; continue; }
    out.push(ch); i++;
  }
  return out;
}
/** Move the relative references of a formula by (dRow, dCol), as filling it across a range does. */
export function shiftFormula(f: string, dRow: number, dCol: number): string {
  return f.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/g, (_, ca, col, ra, row) => `${ca}${ca ? col : colName(colNum(col) + dCol)}${ra}${ra ? row : Number(row) + dRow}`);
}

/* ------------------------------- tools -------------------------------- */

const csv = (rows: unknown[][]) => rows.map(r => r.map(v => { const s = v === null ? "" : String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }).join(",")).join("\n");
const display = (v: ReturnType<Workbook["value"]>) => v;

export function runTool(wb: Workbook, tool: string, a: Record<string, any>): ToolResultLike {
  try { return { ok: true, data: handle(wb, tool, a) }; }
  catch (e) { if (!(e instanceof ToolError)) return { ok: false, error: { code: "invalid_arguments", message: String((e as Error).message) } }; if (e instanceof ToolError) return { ok: false, error: { code: e.code, message: e.message, ...(e.detail !== undefined ? { detail: e.detail } : {}) } }; throw e; }
}

/** Whole-column ("C:C") and whole-row ("3:5") references, which Excel accepts, limited to the used range. */
function normAddr(wb: Workbook, s: Sheet, a: string): string {
  const x = a.replace(/$/g, ""), u = wb.used(s) ?? { c2: 1, r2: 1 };
  let m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(x); if (m) return `${m[1]}1:${m[2]}${u.r2}`;
  m = /^(d+):(d+)$/.exec(x); if (m) return `A${m[1]}:${colName(u.c2)}${m[2]}`;
  return a;
}
function readGrid(wb: Workbook, s: Sheet, address: string) {
  const { c1, r1, c2, r2 } = parseRange(normAddr(wb, s, address)); const values: unknown[][] = [], formulas: unknown[][] = [];
  for (let r = r1; r <= r2; r++) { const v: unknown[] = [], f: unknown[] = []; for (let c = c1; c <= c2; c++) { const k = key(c, r); v.push(display(wb.value(s, k))); const cell = s.cells.get(k); f.push(cell?.f !== undefined ? (cell.f.startsWith("=") ? cell.f : `=${cell.f}`) : display(wb.value(s, k))); } values.push(v); formulas.push(f); }
  return { values, formulas, address: `${colName(c1)}${r1}:${colName(c2)}${r2}`, rows: r2 - r1 + 1, columns: c2 - c1 + 1 };
}

function handle(wb: Workbook, tool: string, a: Record<string, any>): unknown {
  switch (tool) {
    case "get_selection": {
      if (!wb.selection) return { workbookName: wb.name, sheet: wb.sheets[wb.active]!.name, address: "A1", rowCount: 1, columnCount: 1, values: [[null]], headers: null, truncated: false };
      const s = wb.sheet(wb.selection.sheet), g = readGrid(wb, s, wb.selection.address);
      return { workbookName: wb.name, sheet: s.name, address: `${s.name}!${g.address}`, rowCount: g.rows, columnCount: g.columns, values: g.values, headers: null, truncated: false };
    }
    case "get_sheets": return { sheets: wb.sheets.map((s, i) => { const u = wb.used(s); return { name: s.name, visibility: "visible", active: i === wb.active, usedRange: u ? `A1:${colName(u.c2)}${u.r2}` : null, tables: [] }; }) };
    case "read_range": case "get_cell_ranges": {
      const s = wb.sheet(a.sheet); const addrs: string[] = tool === "read_range" ? [a.address] : (a.ranges ?? [a.address]).map((r: any) => (typeof r === "string" ? r : r.address));
      const parts = addrs.map(ad => { const g = readGrid(wb, s, ad); return a.format === "csv" ? { sheet: s.name, address: g.address, csv: csv(g.values) } : { sheet: s.name, address: g.address, values: g.values, ...(a.includeFormulas || tool === "get_cell_ranges" ? { formulas: g.formulas } : {}) }; });
      return tool === "read_range" ? parts[0] : { ranges: parts };
    }
    case "search_data": {
      const hits: unknown[] = []; const term = String(a.searchTerm); const re = a.useRegex ? new RegExp(term, a.matchCase ? "" : "i") : null;
      for (const s of a.sheet ? [wb.sheet(a.sheet)] : wb.sheets) for (const [k, c] of s.cells) { const v = a.matchFormulas ? (c.f ?? "") : String(wb.value(s, k) ?? ""); const ok = re ? re.test(v) : a.matchCase ? v.includes(term) : v.toLowerCase().includes(term.toLowerCase()); if (ok) hits.push({ sheet: s.name, address: k, value: wb.value(s, k) }); }
      return { matches: hits.slice(0, a.limit ?? 50), total: hits.length };
    }
    case "write_range": {
      const s = wb.sheet(a.sheet); const p = parseRange(a.address); const rows = a.values as unknown[][]; const w = Math.max(...rows.map(r => r.length));
      const single = p.c1 === p.c2 && p.r1 === p.r2;
      if (a.mode !== "top_left" && !single && (p.r2 - p.r1 + 1 !== rows.length || p.c2 - p.c1 + 1 !== w)) throw new ToolError("size_mismatch", `address ${a.address} is ${p.r2 - p.r1 + 1}x${p.c2 - p.c1 + 1} but values is ${rows.length}x${w}`);
      const occupied: string[] = [];
      rows.forEach((row, dr) => row.forEach((_v, dc) => { const k = key(p.c1 + dc, p.r1 + dr); if (!wb.isEmpty(s, k)) occupied.push(k); }));
      if (occupied.length && !a.allowOverwrite) throw new ToolError("would_overwrite", `Refused: ${occupied.length} target cell(s) already contain data (${occupied.slice(0, 20).join(", ")}). Read them and ask the user, then set allowOverwrite=true only after they agree or asked to replace data.`, { cells: occupied.slice(0, 20) });
      rows.forEach((row, dr) => row.forEach((v, dc) => { const k = key(p.c1 + dc, p.r1 + dr); if (typeof v === "string" && v.startsWith("=")) s.cells.set(k, { v: null, f: v }); else s.cells.set(k, { v: v as never }); }));
      return { sheet: s.name, address: `${colName(p.c1)}${p.r1}:${colName(p.c1 + w - 1)}${p.r1 + rows.length - 1}`, rowsWritten: rows.length, columnsWritten: w };
    }
    case "set_formula": {
      const s = wb.sheet(a.sheet); const p = parseRange(a.address); const occupied: string[] = [];
      for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) if (!wb.isEmpty(s, key(c, r))) occupied.push(key(c, r));
      if (occupied.length && !a.allowOverwrite) throw new ToolError("would_overwrite", `Refused: ${occupied.length} target cell(s) already contain data (${occupied.slice(0, 20).join(", ")}).`, { cells: occupied.slice(0, 20) });
      const base = String(a.formula).startsWith("=") ? String(a.formula) : `=${a.formula}`;
      for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) s.cells.set(key(c, r), { v: null, f: shiftFormula(base, r - p.r1, c - p.c1) });
      const g = readGrid(wb, s, a.address);
      return { sheet: s.name, address: g.address, cellsWritten: g.rows * g.columns, formulaResults: g.values };
    }
    case "clear_range": { const s = wb.sheet(a.sheet); const p = parseRange(a.address); for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) s.cells.delete(key(c, r)); return { sheet: s.name, address: a.address, cleared: true }; }
    case "insert_rows": {
      const s = wb.sheet(a.sheet); const at = parseCell(/^\d+$/.test(String(a.at)) ? `A${a.at}` : String(a.at)).row; const count = a.count as number;
      const first = (a.position ?? "above") === "below" ? at + 1 : at; const moved = new Map<string, Cell>();
      for (const [k, c] of s.cells) { const p = parseCell(k); moved.set(p.row >= first ? key(p.col, p.row + count) : k, c); }
      s.cells = moved; return { sheet: s.name, insertedAt: `${first}:${first + count - 1}`, rowsInserted: count };
    }
    case "modify_structure": {
      const s = wb.sheet(a.sheet); const dim = a.dimension; const ref = String(a.reference ?? ""); const n = (a.count ?? 1) as number;
      if (a.operation === "delete" && dim === "rows") { const at = Number(ref); const moved = new Map<string, Cell>(); for (const [k, c] of s.cells) { const p = parseCell(k); if (p.row >= at && p.row < at + n) continue; moved.set(p.row >= at + n ? key(p.col, p.row - n) : k, c); } s.cells = moved; return { operation: "delete", dimension: "rows", reference: ref, count: n }; }
      if (a.operation === "hide") { if (dim === "rows") for (let i = 0; i < n; i++) s.hiddenRows.add(Number(ref) + i); else for (let i = 0; i < n; i++) s.hiddenCols.add(colNum(ref) + i); return { operation: "hide", reference: ref }; }
      if (a.operation === "delete" && dim === "columns") { const at = colNum(ref); const moved = new Map<string, Cell>(); for (const [k, c] of s.cells) { const p = parseCell(k); if (p.col >= at && p.col < at + n) continue; moved.set(p.col >= at + n ? key(p.col - n, p.row) : k, c); } s.cells = moved; return { operation: "delete", dimension: "columns", reference: ref, count: n }; }
      if (a.operation === "insert" && dim === "rows") return handle(wb, "insert_rows", { sheet: a.sheet, at: ref, count: n, position: a.position === "after" ? "below" : "above" });
      return { operation: a.operation, note: "accepted" };
    }
    case "manage_sheet": {
      if (a.operation === "add") { const nm = a.name ?? `Sheet${wb.sheets.length + 1}`; if (wb.sheets.some(x => x.name.toLowerCase() === nm.toLowerCase())) throw new ToolError("already_exists", `A sheet named '${nm}' already exists.`); wb.addSheet(nm); wb.active = wb.sheets.length - 1; return { name: nm, added: true }; }
      if (a.operation === "activate") { wb.active = wb.sheets.indexOf(wb.sheet(a.sheet)); return { activated: wb.sheets[wb.active]!.name }; }
      if (a.operation === "rename") { wb.sheet(a.sheet).name = a.newName; return { renamed: a.newName }; }
      return { operation: a.operation, note: "accepted" };
    }
    case "find_errors": {
      const out: unknown[] = [];
      for (const s of a.sheet ? [wb.sheet(a.sheet)] : wb.sheets) for (const k of s.cells.keys()) { const v = wb.value(s, k); if (typeof v === "string" && ERR.test(v)) out.push({ sheet: s.name, address: k, error: v, formula: s.cells.get(k)!.f ?? null }); }
      return { total: out.length, errors: out };
    }
    case "get_objects": return { charts: wb.charts, shapes: [], tables: [], pivotTables: [] };
    case "check_formula_consistency": return { total: 0, issues: [] };
    case "modify_chart": { wb.charts.push({ ...a }); return { operation: a.operation, name: a.name ?? "Chart 1", created: true }; }
    case "format_range": case "resize_range": case "conditional_format": case "page_layout": case "merge_range": return { accepted: true };
    case "set_cells": throw new ToolError("unsupported", "set_cells is not simulated in this eval; use write_range.");
    default: wb.unsimulated.push(tool); throw new ToolError("unsupported", `${tool} is not available in this test workbook.`);
  }
}
