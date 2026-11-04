import { parseCell, type SheetRef } from "../../../protocol/src/index.js";

/**
 * What a workbook-changing tool may touch, worked out from its arguments alone (no Excel needed), so the journal
 * can photograph exactly that before the tool runs. A plan never over-promises: tools it cannot reverse say so.
 */

export interface RangeTarget {
  sheet?: SheetRef;
  address: string;
  /** Use only the top-left cell of `address` and grow it to this many rows/columns (write_range). */
  grow?: { rows: number; columns: number };
}

/** Every visible sheet's used range (find_replace without a sheet). */
export interface AllSheetsTarget { allSheets: true }

export type UndoTarget = RangeTarget | AllSheetsTarget;

export interface LayoutSpec {
  sheet?: SheetRef;
  /** Inclusive 1-based column/row spans whose width/height the tool may change. */
  columns?: [number, number];
  rows?: [number, number];
  /** The whole address given (e.g. `A:D`, `3:3`, a cell range) when the spans are only known from Excel. */
  address?: string;
  dimensions: Array<"width" | "height">;
}

export type UndoPlan =
  | {
      kind: "cells";
      targets: UndoTarget[];
      /** Also photograph fills, fonts, borders and alignment (slower, so only for tools that change them). */
      style: boolean;
      layout?: LayoutSpec;
      /** Merged areas inside the targets are photographed and re-created. */
      merged?: boolean;
      /** Things this tool also changes that the journal does not restore; shown next to the undo button. */
      caveats?: string[];
    }
  | { kind: "insert_rows"; sheet?: SheetRef; row: number; count: number }
  /** The call cannot change the workbook (a preview, a no-op). */
  | { kind: "none" }
  | { kind: "irreversible"; reason: string };

const IRREVERSIBLE: Record<string, string> = {
  execute_office_js: "free-form Office.js can change anything, so there is no record of what to put back",
  modify_structure: "inserting, deleting, hiding or grouping rows and columns is not recorded",
  manage_sheet: "sheet operations are not recorded",
  insert_sheets_from_file: "inserted sheets are not recorded",
  manage_names: "named ranges are not recorded",
  page_layout: "print layout is not recorded",
  modify_chart: "charts are not recorded",
  modify_pivot_table: "pivot tables are not recorded",
  modify_table: "tables are not recorded",
  add_shape: "shapes are not recorded",
  add_connector: "connectors are not recorded",
  add_line: "lines are not recorded",
  update_shape: "shapes are not recorded",
  delete_shapes: "deleted shapes are not recorded",
  filter_range: "filters are not recorded",
  conditional_format: "conditional formats are not recorded",
  data_validation: "data validation is not recorded",
  record_test_results: "test results written to the sheet are not recorded",
};

type Args = Record<string, any>;

/** Bounding box of cell addresses such as `B2` or `$C$5`; null when one is not a plain cell. */
function boundingBox(cells: string[]): { top: number; left: number; bottom: number; right: number } | null {
  let box: { top: number; left: number; bottom: number; right: number } | null = null;
  for (const address of cells) {
    const cell = parseCell(address);
    if (!cell) return null;
    box = box
      ? { top: Math.min(box.top, cell.row), left: Math.min(box.left, cell.col), bottom: Math.max(box.bottom, cell.row), right: Math.max(box.right, cell.col) }
      : { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };
  }
  return box;
}

function toA1(box: { top: number; left: number; bottom: number; right: number }): string {
  const col = (n: number) => {
    let out = "";
    for (let v = n; v > 0; v = Math.floor((v - 1) / 26)) out = String.fromCharCode(65 + ((v - 1) % 26)) + out;
    return out;
  };
  return `${col(box.left)}${box.top}:${col(box.right)}${box.bottom}`;
}

export function planUndo(tool: string, input: unknown): UndoPlan {
  const args = (input && typeof input === "object" ? input : {}) as Args;
  const sheet = args.sheet as SheetRef | undefined;

  switch (tool) {
    case "write_range": {
      const values = Array.isArray(args.values) ? (args.values as unknown[][]) : [];
      const rows = values.length;
      const columns = values.reduce<number>((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
      if (typeof args.address !== "string" || rows === 0 || columns === 0) return { kind: "none" };
      const notes = args.notes && typeof args.notes === "object" && Object.keys(args.notes).length > 0;
      return {
        kind: "cells",
        targets: [{ sheet, address: args.address, grow: { rows, columns } }],
        style: false,
        ...(notes ? { caveats: ["cell notes are not restored"] } : {}),
      };
    }
    case "set_formula":
      return typeof args.address === "string" ? { kind: "cells", targets: [{ sheet, address: args.address }], style: false } : { kind: "none" };
    case "clear_range": {
      if (typeof args.address !== "string") return { kind: "none" };
      const clearType = args.clearType ?? "contents";
      return {
        kind: "cells",
        targets: [{ sheet, address: args.address }],
        style: clearType !== "contents",
        merged: clearType !== "contents",
        ...(clearType === "all"
          ? { caveats: ["notes, hyperlinks, data validation and conditional formats on these cells are not restored"] }
          : clearType === "formats" ? { caveats: ["conditional formats on these cells are not restored"] } : {}),
      };
    }
    case "format_range": {
      if (typeof args.address !== "string") return { kind: "none" };
      const format = (args.format ?? {}) as Args;
      const sizing = format.columnWidth !== undefined || format.rowHeight !== undefined || format.autofitColumns;
      return {
        kind: "cells",
        targets: [{ sheet, address: args.address }],
        style: true,
        ...(sizing
          ? {
              layout: {
                sheet,
                address: args.address,
                dimensions: [
                  ...(format.columnWidth !== undefined || format.autofitColumns ? (["width"] as const) : []),
                  ...(format.rowHeight !== undefined ? (["height"] as const) : []),
                ],
              },
            }
          : {}),
      };
    }
    case "find_replace":
      if (!args.apply) return { kind: "none" };
      return {
        kind: "cells",
        targets: typeof args.address === "string" && sheet ? [{ sheet, address: args.address }] : sheet ? [{ sheet, address: "A:XFD" }] : [{ allSheets: true }],
        style: false,
      };
    case "copy_to": {
      if (typeof args.destination !== "string") return { kind: "none" };
      const copyType = args.copyType ?? "all";
      const formats = copyType === "all" || copyType === "formats";
      return {
        kind: "cells",
        targets: [{ sheet, address: args.destination }],
        style: formats,
        merged: formats,
        ...(formats ? { caveats: ["conditional formats and data validation copied from the source are not removed"] } : {}),
      };
    }
    case "sort_range":
      return typeof args.address === "string" ? { kind: "cells", targets: [{ sheet, address: args.address }], style: true } : { kind: "none" };
    case "merge_range":
      return typeof args.address === "string" ? { kind: "cells", targets: [{ sheet, address: args.address }], style: true, merged: true } : { kind: "none" };
    case "resize_range": {
      const dimensions: Array<"width" | "height"> = [];
      if (args.width) dimensions.push("width");
      if (args.height) dimensions.push("height");
      if (dimensions.length === 0) return { kind: "none" };
      // Without an address the whole sheet is resized, including columns nobody has used.
      if (typeof args.address !== "string") return { kind: "irreversible", reason: "resizing a whole sheet is not recorded" };
      return { kind: "cells", targets: [], style: false, layout: { sheet, address: args.address, dimensions } };
    }
    case "set_cells": {
      const cells = args.cells && typeof args.cells === "object" ? Object.keys(args.cells as object) : [];
      const box = boundingBox(cells);
      if (!box) return { kind: "irreversible", reason: "the cell addresses could not be read" };
      const specs = Object.values(args.cells as Record<string, Args>);
      const targets: UndoTarget[] = [{ sheet, address: toA1(box) }];
      if (typeof args.copyTo === "string") targets.push({ sheet, address: args.copyTo });
      const dimensions: Array<"width" | "height"> = [];
      if (args.resizeWidth) dimensions.push("width");
      if (args.resizeHeight) dimensions.push("height");
      return {
        kind: "cells",
        targets,
        style: specs.some((spec) => spec?.style !== undefined || spec?.borders !== undefined) || typeof args.copyTo === "string",
        merged: typeof args.copyTo === "string",
        ...(dimensions.length ? { layout: { sheet, columns: [box.left, box.right] as [number, number], rows: [box.top, box.bottom] as [number, number], dimensions } } : {}),
        ...(() => {
          const caveats = [
            ...(specs.some((spec) => spec?.note !== undefined) ? ["cell notes are not restored"] : []),
            ...(typeof args.copyTo === "string" ? ["conditional formats and data validation copied along are not removed"] : []),
          ];
          return caveats.length ? { caveats } : {};
        })(),
      };
    }
    case "insert_rows": {
      const at = typeof args.at === "string" ? /(\d+)$/.exec(args.at) : null;
      if (!at) return { kind: "irreversible", reason: "the row to insert at could not be read" };
      const row = Number.parseInt(at[1]!, 10) + (args.position === "below" ? 1 : 0);
      return { kind: "insert_rows", sheet, row, count: Number(args.count) || 1 };
    }
    default:
      return tool in IRREVERSIBLE ? { kind: "irreversible", reason: IRREVERSIBLE[tool]! } : { kind: "irreversible", reason: "this tool is not recorded" };
  }
}
