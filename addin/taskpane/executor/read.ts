import {
  EXECUTOR_CAPS,
  numberToColumn,
  type CellValue,
  type ReadRangeInput,
  type SearchDataInput,
} from "../../../protocol/src/index.js";
import {
  normalizeValues,
  resolveSheet,
  stripSheet,
  toCsv,
  truncateCells,
  workbookName,
} from "./common.js";

import { readActiveShape } from "./shapes.js";

/** Office.js cannot return a range while a shape, chart or other object is selected. */
export const OBJECT_SELECTED_NOTE =
  "A shape, chart or other object is selected, not cells. Use list_shapes / get_objects to see what is on the sheet, " +
  "or ask the user to select the cells they mean.";

/** Small identity read for the composer; never fetch cell values or shape text. */
export async function getSelectionLabel(): Promise<{ label: string; title: string }> {
  return Excel.run(async context => {
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.19")) {
      const shape = context.workbook.getActiveShapeOrNullObject();
      // A null active-shape proxy must be checked before asking for identity.
      // Reading name/type on it can trigger Excel's "Reference isn't valid" dialog.
      shape.load("isNullObject");
      await context.sync();
      if (!shape.isNullObject) {
        shape.load("id,name,type");
        await context.sync();
        return { label: shape.name, title: `Selected ${shape.type === "Line" ? "line / connector" : "shape"}: ${shape.name} (${shape.id})` };
      }
    }
    const range = context.workbook.getSelectedRange();
    range.load("address,rowCount,columnCount");
    await context.sync();
    const cells = range.rowCount * range.columnCount;
    return {
      label: cells > 1 ? `${range.address} · ${cells} cells` : range.address,
      title: "Current selection (sent with your message)",
    };
  });
}

export async function getSelection(): Promise<unknown> {
  try {
    return await readSelection();
  } catch (error) {
    if ((error as { code?: string })?.code !== "InvalidSelection") throw error;
    return Excel.run(async (context) => {
      const selected = await readActiveShape(context);
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,id");
      await context.sync();
      return {
        workbookName: workbookName(),
        workbookUrl: Office.context.document.url,
        sheet: sheet.name,
        sheetId: sheet.id,
        address: "",
        rowCount: 0,
        columnCount: 0,
        values: [],
        headers: null,
        truncated: false,
        ...selected,
        note: selected.shapeSelectionStatus === "unavailable"
          ? OBJECT_SELECTED_NOTE + " Identifying the active shape requires ExcelApi 1.19; this host does not support it."
          : OBJECT_SELECTED_NOTE,
      };
    });
  }
}

async function readSelection(): Promise<unknown> {
  return Excel.run(async (context) => {
    const selected = await readActiveShape(context);
    if (selected.activeShape) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name,id");
      await context.sync();
      return {
        workbookName: workbookName(), workbookUrl: Office.context.document.url, sheet: sheet.name, sheetId: sheet.id,
        address: "", rowCount: 0, columnCount: 0, values: [], headers: null, truncated: false,
        ...selected,
        note: "Active shape identified. This is the active shape, not an enumeration of a multi-selection. Use list_shapes (including groups and connections) to inspect the entire flow.",
      };
    }
    const range = context.workbook.getSelectedRange();
    range.load("address,rowCount,columnCount,rowIndex,columnIndex,values");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name,id");
    await context.sync();

    const truncated = range.rowCount * range.columnCount > EXECUTOR_CAPS.selectionCells;
    const values = normalizeValues(range.values);
    return {
      workbookName: workbookName(),
      workbookUrl: Office.context.document.url,
      sheet: sheet.name,
      sheetId: sheet.id,
      address: stripSheet(range.address),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values: truncated ? truncateCells(values, EXECUTOR_CAPS.selectionCells) : values,
      headers: await detectHeaders(context, sheet, range),
      truncated,
    };
  });
}

/**
 * Column headers for the selected columns: the header row of the Excel Table
 * the selection touches, otherwise the first row of the used range when the
 * selection starts below it. Null when the selection is itself the header row.
 */
async function detectHeaders(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  selection: Excel.Range,
): Promise<string[] | null> {
  let headerRow: number | null = null;

  sheet.tables.load("items");
  await context.sync();
  const hits = sheet.tables.items.map((table) => {
    const intersection = table.getRange().getIntersectionOrNullObject(selection);
    intersection.load("isNullObject");
    const header = table.getHeaderRowRange();
    header.load("rowIndex");
    return { intersection, header };
  });
  if (hits.length > 0) await context.sync();
  const hit = hits.find((h) => !h.intersection.isNullObject);
  if (hit) headerRow = hit.header.rowIndex;

  if (headerRow === null) {
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("isNullObject,rowIndex");
    await context.sync();
    if (!used.isNullObject) headerRow = used.rowIndex;
  }
  if (headerRow === null || selection.rowIndex <= headerRow) return null;

  const header = sheet.getRangeByIndexes(headerRow, selection.columnIndex, 1, selection.columnCount);
  header.load("values");
  await context.sync();
  const cells = (header.values[0] ?? []).map((v) => (v === null || v === undefined ? "" : String(v)));
  return cells.some((c) => c !== "") ? cells : null;
}

export async function getSheets(): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/id,items/visibility");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    await context.sync();

    const usedRanges = sheets.items.map((sheet) => sheet.getUsedRangeOrNullObject(true));
    usedRanges.forEach((range) => range.load("address"));
    await context.sync();

    return {
      activeSheet: active.name,
      sheets: sheets.items.map((sheet, index) => ({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility,
        usedRange: usedRanges[index]!.isNullObject ? null : stripSheet(usedRanges[index]!.address),
        isActive: sheet.name === active.name,
      })),
    };
  });
}

export async function readRange(input: ReadRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const full = sheet.getRange(input.address);
    full.load("address,rowCount,columnCount,rowIndex,columnIndex");
    await context.sync();

    // Page by rows: bounded by maxRows and by the executor's cell cap.
    const cap = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / Math.max(1, full.columnCount)));
    const rows = Math.max(0, Math.min(input.maxRows ?? cap, cap, full.rowCount - input.offset));
    const empty = rows === 0;
    const page = empty ? null : sheet.getRangeByIndexes(full.rowIndex + input.offset, full.columnIndex, rows, full.columnCount);
    if (page) {
      page.load("values");
      if (input.includeFormulas) page.load("formulas");
      await context.sync();
    }

    const values = page ? normalizeValues(page.values) : [];
    const end = input.offset + rows;
    const hasMore = end < full.rowCount;
    const base = {
      sheet: sheet.name,
      address: stripSheet(full.address),
      rowCount: full.rowCount,
      columnCount: full.columnCount,
      offset: input.offset,
      returnedRows: rows,
      hasMore,
      nextOffset: hasMore ? end : null,
      truncated: hasMore,
    };
    const formulas = input.includeFormulas && page ? (page.formulas as string[][]) : input.includeFormulas ? [] : null;
    return input.format === "csv" ? { ...base, csv: toCsv(values), formulas } : { ...base, values, formulas };
  });
}

interface Match {
  sheet: string;
  address: string;
  value: CellValue;
}

export async function searchData(input: SearchDataInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();

    const targets = input.sheet
      ? [resolveSheet(context, input.sheet)]
      : sheets.items.filter((sheet) => sheet.visibility === "Visible");
    if (input.sheet) targets[0]!.load("name");

    const needle = input.matchCase ? input.searchTerm : input.searchTerm.toLowerCase();
    const wanted = input.offset + input.limit + 1; // one extra tells us whether more exist
    const matches: Match[] = [];
    let scanned = 0;
    let scanTruncated = false;

    outer: for (const sheet of targets) {
      let used = sheet.getUsedRangeOrNullObject(true);
      used.load("isNullObject");
      await context.sync();
      if (used.isNullObject) continue;
      if (input.address) {
        // Intersect only once we know the used range exists (null objects cannot be chained).
        used = used.getIntersectionOrNullObject(sheet.getRange(input.address));
      }
      used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      if (used.isNullObject) continue;

      // Read in row chunks so one call never loads an unbounded block.
      const rowsPerChunk = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / used.columnCount));
      for (let offset = 0; offset < used.rowCount; offset += rowsPerChunk) {
        if (scanned >= EXECUTOR_CAPS.searchScanCells) {
          scanTruncated = true;
          break outer;
        }
        const rows = Math.min(rowsPerChunk, used.rowCount - offset);
        const chunk = sheet.getRangeByIndexes(used.rowIndex + offset, used.columnIndex, rows, used.columnCount);
        chunk.load("values");
        await context.sync();
        scanned += rows * used.columnCount;

        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < used.columnCount; c++) {
            const value = chunk.values[r]?.[c];
            if (value === null || value === undefined || value === "") continue;
            const text = input.matchCase ? String(value) : String(value).toLowerCase();
            if (input.wholeCell ? text === needle : text.includes(needle)) {
              matches.push({
                sheet: sheet.name,
                address: `${numberToColumn(used.columnIndex + c + 1)}${used.rowIndex + offset + r + 1}`,
                value: value as CellValue,
              });
              if (matches.length >= wanted) break outer;
            }
          }
        }
      }
    }

    const page = matches.slice(input.offset, input.offset + input.limit);
    const hasMore = matches.length > input.offset + input.limit;
    return {
      matches: page,
      totalFound: matches.length,
      returned: page.length,
      hasMore,
      nextOffset: hasMore ? input.offset + input.limit : null,
      scanTruncated,
    };
  });
}
