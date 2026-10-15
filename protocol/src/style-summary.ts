import { numberToColumn } from "./address.js";

/** Compact, readable style of a cell (defaults omitted); shared by the live reader and the file reader. */
export interface StyleOut {
  bold?: true;
  italic?: true;
  underline?: string;
  strikethrough?: true;
  fontSize?: number;
  fontColor?: string;
  fontFamily?: string;
  backgroundColor?: string;
  horizontalAlignment?: string;
  wrapText?: true;
  numberFormat?: string;
}

export type BorderOut = Partial<Record<"top" | "bottom" | "left" | "right", string>>;

const OFFICE_BORDER_STYLE: Record<string, string> = { Continuous: "solid", Dash: "dashed", Dot: "dotted", Double: "double" };

/** `thin solid #000000`, or null for no border. */
export function describeBorder(side?: { style?: string; color?: string; weight?: string }): string | null {
  if (!side || !side.style || side.style === "None") return null;
  const weight = (side.weight ?? "Thin").toLowerCase();
  const style = OFFICE_BORDER_STYLE[side.style] ?? side.style.toLowerCase();
  return `${weight === "hairline" ? "thin" : weight} ${style} ${(side.color ?? "#000000").toLowerCase()}`;
}

/**
 * Group identical per-cell styles into rectangles: runs along each row, then runs
 * with the same columns on consecutive rows are merged. Empty styles are skipped.
 */
export function groupByRectangles<T extends object>(grid: Array<Array<T | null>>, firstRow: number, firstCol: number): Record<string, T> {
  type Open = { r0: number; r1: number; c0: number; c1: number; key: string; value: T };
  const done: Open[] = [];
  let open = new Map<string, Open>();

  grid.forEach((row, r) => {
    const next = new Map<string, Open>();
    let c = 0;
    while (c < row.length) {
      const cell = row[c];
      const key = cell && Object.keys(cell).length > 0 ? JSON.stringify(cell) : null;
      if (!key) {
        c++;
        continue;
      }
      let end = c;
      while (end + 1 < row.length && JSON.stringify(row[end + 1]) === key) end++;
      const rectKey = `${c}:${end}:${key}`;
      const above = open.get(rectKey);
      if (above) {
        above.r1 = r;
        next.set(rectKey, above);
        open.delete(rectKey);
      } else {
        next.set(rectKey, { r0: r, r1: r, c0: c, c1: end, key, value: cell as T });
      }
      c = end + 1;
    }
    done.push(...open.values());
    open = next;
  });
  done.push(...open.values());

  const out: Record<string, T> = {};
  for (const rect of done.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0)) {
    const a1 = (r: number, c: number) => `${numberToColumn(firstCol + c + 1)}${firstRow + r + 1}`;
    const from = a1(rect.r0, rect.c0);
    const to = a1(rect.r1, rect.c1);
    out[from === to ? from : `${from}:${to}`] = rect.value;
  }
  return out;
}

