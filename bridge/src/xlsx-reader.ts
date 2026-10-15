import { unzipSync, strFromU8 } from "fflate";
import { XMLParser } from "fast-xml-parser";
import {
  columnToNumber,
  describeBorder,
  groupByRectangles,
  numberToColumn,
  parseRange,
  type BorderOut,
  type StyleOut,
} from "../../protocol/src/index.js";

/**
 * Read-only inspection of a CLOSED .xlsx: sheets, cells, merges, sizes, styles and drawn shapes. Used so the
 * assistant can learn from template files without opening them in Excel. Pure (bytes in, JSON out).
 */

const ARRAY_TAGS = new Set([
  "sheet", "c", "row", "mergeCell", "si", "r", "xf", "font", "fill", "border", "numFmt", "definedName", "col", "Relationship",
  "dataValidation", "twoCellAnchor", "oneCellAnchor", "absoluteAnchor", "sp", "cxnSp", "grpSp", "pic", "p", "a:p", "a:r", "xdr:sp",
]);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  removeNSPrefix: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  isArray: (name) => ARRAY_TAGS.has(name),
});

type Node = Record<string, any>;
const arr = <T = Node>(v: unknown): T[] => (v === undefined || v === null ? [] : Array.isArray(v) ? (v as T[]) : [v as T]);
const text = (v: unknown): string => (typeof v === "string" ? v : v && typeof v === "object" ? String((v as Node)["#text"] ?? "") : v === undefined ? "" : String(v));

export class XlsxError extends Error {}

interface Package {
  files: Record<string, Uint8Array>;
  xml(path: string): Node | null;
}

function open(bytes: Uint8Array): Package {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes, { filter: (f) => !f.name.startsWith("xl/media/") && !f.name.includes("printerSettings") });
  } catch {
    throw new XlsxError("not a valid .xlsx file (zip could not be read)");
  }
  return {
    files,
    xml(path) {
      const raw = files[path];
      return raw ? (parser.parse(strFromU8(raw)) as Node) : null;
    },
  };
}

/* -------------------------------------------------------------------------- */
/* Colours                                                                     */
/* -------------------------------------------------------------------------- */

const THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

function themeColors(pkg: Package): string[] {
  const scheme = pkg.xml("xl/theme/theme1.xml")?.theme?.themeElements?.clrScheme ?? {};
  return THEME_ORDER.map((key) => {
    const entry = scheme[key] ?? {};
    return ("#" + (entry.srgbClr?.["@_val"] ?? entry.sysClr?.["@_lastClr"] ?? "000000")).toUpperCase();
  });
}

function hexToHls(hex: string): [number, number, number] {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? ((g - b) / d + (g < b ? 6 : 0)) / 6 : max === g ? ((b - r) / d + 2) / 6 : ((r - g) / d + 4) / 6;
  return [h, l, s];
}

function hlsToHex(h: number, l: number, s: number): string {
  const f = (p: number, q: number, t0: number) => {
    const t = (t0 + 1) % 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let rgb: number[];
  if (s === 0) rgb = [l, l, l];
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    rgb = [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)];
  }
  return "#" + rgb.map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Excel tint: positive lightens toward white, negative darkens. */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, l, s] = hexToHls(hex);
  return hlsToHex(h, tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint, s);
}

/** SpreadsheetML colour element (`rgb` / `theme`+`tint` / `indexed`) -> `#RRGGBB` or null (automatic). */
function cellColor(el: Node | undefined, theme: string[]): string | null {
  if (!el) return null;
  if (el["@_rgb"]) return ("#" + String(el["@_rgb"]).slice(-6)).toUpperCase();
  if (el["@_theme"] !== undefined) return applyTint(theme[Number(el["@_theme"])] ?? "#000000", Number(el["@_tint"] ?? 0));
  return null;
}

/** DrawingML colour child (`srgbClr` / `schemeClr` with lumMod/lumOff) -> `#RRGGBB`. */
function drawingColor(holder: Node | undefined, theme: string[]): string | null {
  if (!holder) return null;
  if (holder.srgbClr) return ("#" + holder.srgbClr["@_val"]).toUpperCase();
  const scheme = holder.schemeClr;
  if (scheme) {
    const map: Record<string, number> = { bg1: 0, tx1: 1, bg2: 2, tx2: 3, lt1: 0, dk1: 1, lt2: 2, dk2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9 };
    let hex = theme[map[scheme["@_val"]] ?? 1] ?? "#000000";
    const mod = Number(scheme.lumMod?.["@_val"] ?? 100000) / 100000;
    const off = Number(scheme.lumOff?.["@_val"] ?? 0) / 100000;
    if (mod !== 1 || off !== 0) {
      const [h, l, s] = hexToHls(hex);
      hex = hlsToHex(h, Math.min(1, l * mod + off), s);
    }
    return hex;
  }
  if (holder.sysClr) return ("#" + (holder.sysClr["@_lastClr"] ?? "000000")).toUpperCase();
  return null;
}

/* -------------------------------------------------------------------------- */
/* Workbook parts                                                              */
/* -------------------------------------------------------------------------- */

interface SheetRef {
  name: string;
  state: string;
  path: string;
}

function sheetRefs(pkg: Package): SheetRef[] {
  const workbook = pkg.xml("xl/workbook.xml")?.workbook;
  if (!workbook) throw new XlsxError("workbook.xml is missing");
  const rels = new Map<string, string>();
  for (const r of arr(pkg.xml("xl/_rels/workbook.xml.rels")?.Relationships?.Relationship)) rels.set(r["@_Id"], r["@_Target"]);
  return arr(workbook.sheets?.sheet).map((s) => {
    const target = String(rels.get(s["@_id"]) ?? "");
    const path = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    return { name: String(s["@_name"]), state: String(s["@_state"] ?? "visible"), path };
  });
}

function sharedStrings(pkg: Package): string[] {
  return arr(pkg.xml("xl/sharedStrings.xml")?.sst?.si).map((si) =>
    // Visible text only: plain <t> and rich <r><t> runs; phonetic (furigana) runs are skipped.
    arr(si.t).map(text).join("") + arr(si.r).map((r) => arr(r.t).map(text).join("")).join(""),
  );
}

interface Styles {
  /** Per cell-xf index. */
  xfs: Array<{ style: StyleOut; borders: BorderOut; fill: boolean }>;
}

const NUMFMT_BUILTIN: Record<string, string> = { "0": "General", "14": "m/d/yyyy", "49": "@", "9": "0%", "10": "0.00%", "2": "0.00", "4": "#,##0.00", "3": "#,##0" };

function readStyles(pkg: Package, theme: string[]): Styles {
  const sheet = pkg.xml("xl/styles.xml")?.styleSheet ?? {};
  const fonts = arr(sheet.fonts?.font);
  const fills = arr(sheet.fills?.fill);
  const borders = arr(sheet.borders?.border);
  const numFmts = new Map(arr(sheet.numFmts?.numFmt).map((n) => [String(n["@_numFmtId"]), String(n["@_formatCode"])]));

  const xfs = arr(sheet.cellXfs?.xf).map((xf) => {
    const font = fonts[Number(xf["@_fontId"] ?? 0)] ?? {};
    const style: StyleOut = {};
    if (font.b !== undefined && font.b?.["@_val"] !== "0") style.bold = true;
    if (font.i !== undefined && font.i?.["@_val"] !== "0") style.italic = true;
    if (font.u !== undefined) style.underline = String(font.u?.["@_val"] ?? "single").toLowerCase();
    if (font.strike !== undefined) style.strikethrough = true;
    if (font.sz) style.fontSize = Number(font.sz["@_val"]);
    if (font.name) style.fontFamily = String(font.name["@_val"]);
    const color = cellColor(font.color, theme);
    if (color && color !== "#000000") style.fontColor = color;

    const pattern = fills[Number(xf["@_fillId"] ?? 0)]?.patternFill;
    let filled = false;
    if (pattern && pattern["@_patternType"] === "solid") {
      const bg = cellColor(pattern.fgColor, theme);
      if (bg && bg !== "#FFFFFF") {
        style.backgroundColor = bg;
        filled = true;
      }
    }
    const align = xf.alignment?.["@_horizontal"];
    if (align && align !== "general") style.horizontalAlignment = String(align);
    if (xf.alignment?.["@_wrapText"] === "1") style.wrapText = true;
    const id = String(xf["@_numFmtId"] ?? "0");
    const code = numFmts.get(id) ?? NUMFMT_BUILTIN[id];
    if (code && code !== "General") style.numberFormat = code;

    const border = borders[Number(xf["@_borderId"] ?? 0)] ?? {};
    const out: BorderOut = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const b = border[side];
      if (!b || !b["@_style"]) continue;
      const weight = { thin: "Thin", hair: "Hairline", medium: "Medium", thick: "Thick", double: "Thin" }[String(b["@_style"])] ?? "Thin";
      const line = { double: "Double", dashed: "Dash", dotted: "Dot", hair: "Continuous" }[String(b["@_style"])] ?? "Continuous";
      out[side] = describeBorder({ style: line, weight, color: cellColor(b.color, theme) ?? "#000000" }) ?? undefined;
    }
    return { style, borders: out, fill: filled };
  });
  return { xfs };
}

/* -------------------------------------------------------------------------- */
/* Public API                                                                  */
/* -------------------------------------------------------------------------- */

export interface SheetOverview {
  name: string;
  state: string;
  dimension: string | null;
  mergeCount: number;
  shapeCount: number;
  tabColor: string | null;
  zoom: number | null;
  gridlines: boolean;
  frozenAt: string | null;
  printArea: string | null;
  orientation: string | null;
}

export interface WorkbookOverview {
  sheets: SheetOverview[];
  definedNames: Array<{ name: string; value: string }>;
}

function sheetPartPaths(pkg: Package, sheetPath: string): { drawing: string | null } {
  const relsPath = sheetPath.replace(/worksheets\/(sheet[^/]*)\.xml$/, "worksheets/_rels/$1.xml.rels");
  const rels = arr(pkg.xml(relsPath)?.Relationships?.Relationship);
  const drawing = rels.find((r) => /\/drawing$/.test(String(r["@_Type"])));
  if (!drawing) return { drawing: null };
  const target = String(drawing["@_Target"]);
  const resolved = target.startsWith("/") ? target.slice(1) : new URL(target, `file:///${sheetPath}`).pathname.slice(1);
  return { drawing: resolved };
}

export function readWorkbookOverview(bytes: Uint8Array): WorkbookOverview {
  const pkg = open(bytes);
  const theme = themeColors(pkg);
  const names = arr(pkg.xml("xl/workbook.xml")?.workbook?.definedNames?.definedName).map((d) => ({
    name: String(d["@_name"]) + (d["@_localSheetId"] !== undefined ? ` [sheet ${d["@_localSheetId"]}]` : ""),
    value: text(d),
  }));
  const refs = sheetRefs(pkg);
  const sheets = refs.map((ref, index): SheetOverview => {
    const ws = pkg.xml(ref.path)?.worksheet ?? {};
    const view = arr(ws.sheetViews?.sheetView)[0] ?? {};
    const pane = view.pane;
    const { drawing } = sheetPartPaths(pkg, ref.path);
    const shapeCount = drawing ? countShapes(pkg.xml(drawing)) : 0;
    const printArea = names.find((n) => n.name.startsWith("_xlnm.Print_Area") && n.name.includes(`[sheet ${index}]`));
    return {
      name: ref.name,
      state: ref.state,
      dimension: ws.dimension?.["@_ref"] ?? null,
      mergeCount: arr(ws.mergeCells?.mergeCell).length,
      shapeCount,
      tabColor: cellColor(ws.sheetPr?.tabColor, theme),
      zoom: view["@_zoomScale"] ? Number(view["@_zoomScale"]) : null,
      gridlines: view["@_showGridLines"] !== "0",
      frozenAt: pane && pane["@_state"] === "frozen" ? String(pane["@_topLeftCell"] ?? "") : null,
      printArea: printArea ? printArea.value.replace(/^.*!/, "") : null,
      orientation: ws.pageSetup?.["@_orientation"] ?? null,
    };
  });
  return {
    sheets,
    definedNames: names.map((n) => ({ name: n.name.replace(/\s\[sheet \d+\]$/, ""), value: n.value })).slice(0, 50),
  };
}

function countShapes(drawing: Node | null): number {
  let n = 0;
  const visit = (node: Node) => {
    for (const key of ["sp", "cxnSp", "pic"]) n += arr(node[key]).length;
    for (const g of arr(node.grpSp)) visit(g);
  };
  for (const key of ["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"]) for (const anchor of arr(drawing?.wsDr?.[key])) visit(anchor);
  return n;
}

export interface ShapeSummary {
  name: string;
  kind: string;
  geometry: string | null;
  text: string | null;
  from: string | null;
  to: string | null;
  fill: string | null;
  line: string | null;
  font: string | null;
}

function anchorCell(point: Node | undefined): string | null {
  if (!point) return null;
  return `${numberToColumn(Number(point.col) + 1)}${Number(point.row) + 1}`;
}

function summariseShapes(drawing: Node | null, theme: string[]): ShapeSummary[] {
  const out: ShapeSummary[] = [];
  const visit = (node: Node, from: string | null, to: string | null) => {
    for (const kind of ["sp", "cxnSp", "pic"] as const) {
      for (const el of arr(node[kind])) {
        const nv = el.nvSpPr ?? el.nvCxnSpPr ?? el.nvPicPr ?? {};
        const sp = el.spPr ?? {};
        const solid = sp.solidFill;
        const fill = sp.noFill !== undefined ? "none" : drawingColor(solid, theme);
        const ln = sp.ln;
        const dash = ln?.prstDash?.["@_val"];
        const head = ln?.headEnd?.["@_type"];
        const tail = ln?.tailEnd?.["@_type"];
        const lineColor = ln?.noFill !== undefined ? "none" : drawingColor(ln?.solidFill, theme);
        const widthPt = ln?.["@_w"] ? Math.round((Number(ln["@_w"]) / 12700) * 100) / 100 : null;
        const line = ln
          ? [lineColor ?? "auto", widthPt ? `${widthPt}pt` : null, dash && dash !== "solid" ? dash : null, head && head !== "none" ? `begin:${head}` : null, tail && tail !== "none" ? `end:${tail}` : null]
              .filter(Boolean)
              .join(" ")
          : null;
        const paragraphs = arr(el.txBody?.p).map((p: Node) => arr(p.r).map((r: Node) => text(r.t)).join("")).filter(Boolean);
        const run = arr(arr(el.txBody?.p)[0]?.r)[0]?.rPr ?? arr(el.txBody?.p)[0]?.endParaRPr;
        const font = run
          ? [run.latin?.["@_typeface"], run["@_sz"] ? `${Number(run["@_sz"]) / 100}pt` : null, drawingColor(run.solidFill, theme)].filter(Boolean).join(" ")
          : null;
        out.push({
          name: String(nv.cNvPr?.["@_name"] ?? ""),
          kind: kind === "cxnSp" ? "connector" : kind === "pic" ? "picture" : "shape",
          geometry: sp.prstGeom?.["@_prst"] ?? null,
          text: paragraphs.length ? paragraphs.join(" / ") : null,
          from,
          to,
          fill,
          line,
          font: font || null,
        });
      }
    }
    for (const group of arr(node.grpSp)) visit(group, from, to);
  };
  for (const key of ["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"]) {
    for (const anchor of arr(drawing?.wsDr?.[key])) visit(anchor, anchorCell(anchor.from), anchorCell(anchor.to));
  }
  return out;
}

export interface SheetDetails {
  sheet: string;
  state: string;
  dimension: string | null;
  range: string;
  defaultColumnWidth: number | null;
  defaultRowHeight: number | null;
  columnWidths: Array<{ columns: string; width: number }>;
  rowHeights: Array<{ rows: string; height: number }>;
  merges: string[];
  cells: Record<string, string | number | boolean | [string | number | boolean, string]>;
  styles?: Record<string, StyleOut>;
  borders?: Record<string, BorderOut>;
  baseline?: { fontFamily?: string; fontSize?: number };
  validations: Array<{ range: string; type: string; formula: string }>;
  shapes?: ShapeSummary[];
  truncated: boolean;
}

export interface SheetOptions {
  range?: string;
  includeStyles?: boolean;
  includeShapes?: boolean;
  maxCells?: number;
}

const MAX_GRID_CELLS = 40_000;

export function readSheetDetails(bytes: Uint8Array, sheetName: string, options: SheetOptions = {}): SheetDetails {
  const pkg = open(bytes);
  const refs = sheetRefs(pkg);
  const ref = refs.find((r) => r.name === sheetName) ?? refs.find((r) => r.name.trim().toLowerCase() === sheetName.trim().toLowerCase());
  if (!ref) throw new XlsxError(`sheet '${sheetName}' not found. Sheets: ${refs.map((r) => r.name).join(" | ")}`);

  const theme = themeColors(pkg);
  const strings = sharedStrings(pkg);
  const styles = readStyles(pkg, theme);
  const ws = pkg.xml(ref.path)?.worksheet ?? {};
  const dimension: string | null = ws.dimension?.["@_ref"] ?? null;

  const bounds = options.range ? parseRange(options.range) : dimension ? parseRange(dimension) : { start: { col: 1, row: 1 }, end: { col: 1, row: 1 } };
  if (!bounds) throw new XlsxError(`bad range '${options.range}'`);
  const r0 = Math.min(bounds.start.row, bounds.end.row);
  const r1 = Math.max(bounds.start.row, bounds.end.row);
  const c0 = Math.min(bounds.start.col, bounds.end.col);
  const c1 = Math.max(bounds.start.col, bounds.end.col);
  const rangeText = `${numberToColumn(c0)}${r0}:${numberToColumn(c1)}${r1}`;

  const cells: SheetDetails["cells"] = {};
  const maxCells = options.maxCells ?? 3_000;
  let truncated = false;
  const gridRows = r1 - r0 + 1;
  const gridCols = c1 - c0 + 1;
  const wantGrid = options.includeStyles !== false && gridRows * gridCols <= MAX_GRID_CELLS;
  const styleGrid: Array<Array<StyleOut | null>> = wantGrid ? Array.from({ length: gridRows }, () => Array<StyleOut | null>(gridCols).fill(null)) : [];
  const borderGrid: Array<Array<BorderOut | null>> = wantGrid ? Array.from({ length: gridRows }, () => Array<BorderOut | null>(gridCols).fill(null)) : [];
  const fontStats: Array<{ family?: string; size?: number }> = [];

  for (const row of arr(ws.sheetData?.row)) {
    const rowNumber = Number(row["@_r"]);
    if (rowNumber < r0 || rowNumber > r1) continue;
    for (const c of arr(row.c)) {
      const addr = String(c["@_r"]);
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      const col = columnToNumber(m[1]!);
      if (col < c0 || col > c1) continue;

      const xf = styles.xfs[Number(c["@_s"] ?? 0)];
      if (wantGrid && xf) {
        fontStats.push({ family: xf.style.fontFamily, size: xf.style.fontSize });
        styleGrid[rowNumber - r0]![col - c0] = xf.style;
        borderGrid[rowNumber - r0]![col - c0] = xf.borders;
      }

      let value: string | number | boolean | null = null;
      const t = c["@_t"];
      if (c.v !== undefined) {
        const raw = text(c.v);
        value = t === "s" ? (strings[Number(raw)] ?? "") : t === "b" ? raw === "1" : t === "str" || t === "e" ? raw : raw !== "" && !Number.isNaN(Number(raw)) ? Number(raw) : raw;
      } else if (c.is) {
        value = arr(c.is.t).map(text).join("");
      }
      const formula = c.f !== undefined ? text(c.f) : "";
      if ((value === null || value === "") && !formula) continue;
      if (Object.keys(cells).length >= maxCells) {
        truncated = true;
        continue;
      }
      cells[addr] = formula ? [value ?? "", `=${formula}`] : (value as string | number | boolean);
    }
  }

  const merges = arr(ws.mergeCells?.mergeCell)
    .map((m) => String(m["@_ref"]))
    .filter((ref0) => {
      const p = parseRange(ref0);
      return p && p.start.row <= r1 && p.end.row >= r0 && p.start.col <= c1 && p.end.col >= c0;
    });

  const colRuns = arr(ws.cols?.col).map((c) => ({
    columns: `${numberToColumn(Number(c["@_min"]))}:${numberToColumn(Math.min(Number(c["@_max"]), 16384))}`,
    width: Number(c["@_width"]),
  }));
  const rowRuns = new Map<number, number[]>();
  for (const row of arr(ws.sheetData?.row)) {
    if (row["@_ht"] && row["@_customHeight"] !== "0") {
      const list = rowRuns.get(Number(row["@_ht"])) ?? [];
      list.push(Number(row["@_r"]));
      rowRuns.set(Number(row["@_ht"]), list);
    }
  }

  const details: SheetDetails = {
    sheet: ref.name,
    state: ref.state,
    dimension,
    range: rangeText,
    defaultColumnWidth: ws.sheetFormatPr?.["@_defaultColWidth"] ? Number(ws.sheetFormatPr["@_defaultColWidth"]) : null,
    defaultRowHeight: ws.sheetFormatPr?.["@_defaultRowHeight"] ? Number(ws.sheetFormatPr["@_defaultRowHeight"]) : null,
    columnWidths: colRuns,
    rowHeights: [...rowRuns.entries()].map(([height, rows]) => ({ rows: compressRuns(rows), height })),
    merges: merges.slice(0, 200),
    cells,
    validations: arr(ws.dataValidations?.dataValidation).map((d) => ({
      range: String(d["@_sqref"]),
      type: String(d["@_type"] ?? ""),
      formula: text(d.formula1),
    })),
    truncated,
  };

  if (wantGrid) {
    const family = mode(fontStats.map((f) => f.family));
    const size = mode(fontStats.map((f) => f.size));
    const strip = (s: StyleOut | null): StyleOut | null => {
      if (!s) return null;
      const copy: StyleOut = { ...s };
      if (copy.fontFamily === family) delete copy.fontFamily;
      if (copy.fontSize === size) delete copy.fontSize;
      return copy;
    };
    details.styles = groupByRectangles(styleGrid.map((row) => row.map(strip)), r0 - 1, c0 - 1);
    details.borders = groupByRectangles(borderGrid, r0 - 1, c0 - 1);
    details.baseline = { fontFamily: family, fontSize: size };
  }

  if (options.includeShapes !== false) {
    const { drawing } = sheetPartPaths(pkg, ref.path);
    if (drawing) details.shapes = summariseShapes(pkg.xml(drawing), theme).slice(0, 300);
  }
  return details;
}

function mode<T>(values: Array<T | undefined>): T | undefined {
  const tally = new Map<T, number>();
  for (const v of values) if (v !== undefined) tally.set(v, (tally.get(v) ?? 0) + 1);
  return [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
}

function compressRuns(rows: number[]): string {
  const sorted = [...rows].sort((a, b) => a - b);
  const runs: string[] = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j]! + 1) j++;
    runs.push(i === j ? String(sorted[i]) : `${sorted[i]}-${sorted[j]}`);
    i = j + 1;
  }
  return runs.slice(0, 40).join(",") + (runs.length > 40 ? ",…" : "");
}

export function listSheetNames(bytes: Uint8Array): string[] {
  return sheetRefs(open(bytes)).map((r) => r.name);
}
