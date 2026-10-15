import {
  describeBorder,
  parseCell,
  type BorderOut,
  type BorderSide,
  type Borders,
  type CellStyle,
  type StyleOut,
} from "../../../protocol/src/index.js";

export { describeBorder, groupByRectangles, type BorderOut, type StyleOut } from "../../../protocol/src/index.js";

/* ----- reading: Office.js cell properties -> compact, human-readable style ----- */

/** The shape returned by `Range.getCellProperties` for the properties we ask for. */
export interface CellProps {
  format?: {
    font?: { bold?: boolean; italic?: boolean; color?: string; name?: string; size?: number; underline?: string; strikethrough?: boolean };
    fill?: { color?: string };
    horizontalAlignment?: string;
    borders?: Partial<Record<"top" | "bottom" | "left" | "right", { style?: string; color?: string; weight?: string }>>;
  };
}

/** Fonts/sizes that cover most of the block become the baseline and are left out of per-range styles. */
export interface Baseline {
  fontFamily?: string;
  fontSize?: number;
}

export function modalBaseline(props: CellProps[][]): Baseline {
  const count = <T>(pick: (p: CellProps) => T | undefined) => {
    const tally = new Map<T, number>();
    for (const row of props) for (const p of row) {
      const v = pick(p);
      if (v !== undefined && v !== null) tally.set(v, (tally.get(v) ?? 0) + 1);
    }
    return [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
  };
  return { fontFamily: count((p) => p.format?.font?.name), fontSize: count((p) => p.format?.font?.size) };
}

export function toStyle(props: CellProps, numberFormat: string | undefined, base: Baseline): StyleOut {
  const font = props.format?.font ?? {};
  const out: StyleOut = {};
  if (font.bold) out.bold = true;
  if (font.italic) out.italic = true;
  if (font.underline && font.underline !== "None") out.underline = font.underline.toLowerCase();
  if (font.strikethrough) out.strikethrough = true;
  if (font.size !== undefined && font.size !== base.fontSize) out.fontSize = font.size;
  if (font.name && font.name !== base.fontFamily) out.fontFamily = font.name;
  if (font.color && font.color.toLowerCase() !== "#000000") out.fontColor = font.color.toLowerCase();
  const fill = props.format?.fill?.color;
  if (fill && fill.toLowerCase() !== "#ffffff") out.backgroundColor = fill.toLowerCase();
  const align = props.format?.horizontalAlignment;
  if (align && align !== "General") out.horizontalAlignment = align.toLowerCase();
  if (numberFormat && numberFormat !== "General") out.numberFormat = numberFormat;
  return out;
}

export function toBorders(props: CellProps): BorderOut {
  const out: BorderOut = {};
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const text = describeBorder(props.format?.borders?.[side]);
    if (text) out[side] = text;
  }
  return out;
}

/* ----- writing: our style schema -> Office.js ----- */

const OFFICE_BORDER_LINE: Record<BorderSide["style"], string> = { solid: "Continuous", dashed: "Dash", dotted: "Dot", double: "Double" };
const OFFICE_BORDER_WEIGHT: Record<BorderSide["weight"], string> = { thin: "Thin", medium: "Medium", thick: "Thick" };

export function applyCellStyle(range: Excel.Range, style: CellStyle): void {
  const fmt = range.format;
  if (style.bold !== undefined) fmt.font.bold = style.bold;
  if (style.italic !== undefined) fmt.font.italic = style.italic;
  if (style.underline !== undefined) fmt.font.underline = style.underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
  if (style.strikethrough !== undefined) fmt.font.strikethrough = style.strikethrough;
  if (style.fontSize !== undefined) fmt.font.size = style.fontSize;
  if (style.fontColor) fmt.font.color = style.fontColor;
  if (style.fontFamily) fmt.font.name = style.fontFamily;
  if (style.backgroundColor) fmt.fill.color = style.backgroundColor;
  if (style.horizontalAlignment) fmt.horizontalAlignment = style.horizontalAlignment;
  if (style.wrapText !== undefined) fmt.wrapText = style.wrapText;
}

export function applyBorders(range: Excel.Range, borders: Borders): void {
  const sides: Array<[Excel.BorderIndex, BorderSide | undefined]> = [
    ["EdgeTop" as Excel.BorderIndex, borders.top ?? borders.all],
    ["EdgeBottom" as Excel.BorderIndex, borders.bottom ?? borders.all],
    ["EdgeLeft" as Excel.BorderIndex, borders.left ?? borders.all],
    ["EdgeRight" as Excel.BorderIndex, borders.right ?? borders.all],
  ];
  for (const [index, side] of sides) {
    if (!side) continue;
    const full: BorderSide = { style: side.style ?? "solid", weight: side.weight ?? "thin", color: side.color ?? "#000000" };
    const border = range.format.borders.getItem(index);
    border.style = OFFICE_BORDER_LINE[full.style] as Excel.BorderLineStyle;
    border.weight = OFFICE_BORDER_WEIGHT[full.weight] as Excel.BorderWeight;
    border.color = full.color;
  }
}

/** Strip `$` and validate a single cell address; returns 0-based row/col. */
export function cellPosition(address: string): { row: number; col: number } {
  const cell = parseCell(address.replace(/\$/g, ""));
  if (!cell) throw new Error(`bad cell address '${address}'`);
  return { row: cell.row - 1, col: cell.col - 1 };
}
