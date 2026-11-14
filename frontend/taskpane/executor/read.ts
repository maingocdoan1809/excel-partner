import {
  EXECUTOR_CAPS,
  type CellValue,
  type ReadRangeInput,
  type SearchDataInput,
} from "../../../protocol/src/index.js";
import {
  ToolFailure,
  normalizeValues,
  resolveSheet,
  stripSheet,
  toCsv,
  truncateCells,
  workbookName, runExcel } from "./common.js";
import { cellAddress } from "./audit.js";

import { readActiveShape } from "./shapes.js";

/** Office.js cannot return a range while a shape, chart or other object is selected. */
export const OBJECT_SELECTED_NOTE =
  "A shape, chart or other object is selected, not cells. Use list_shapes / get_objects to see what is on the sheet, " +
  "or ask the user to select the cells they mean.";

/** Small identity read for the composer; never fetch cell values or shape text. */
export async function getSelectionLabel(): Promise<{ label: string; title: string }> {
  // Plain Excel.run: this feeds the selection chip, which must fail fast while the person is typing in a cell.
  return Excel.run(async context => {
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.19")) {
      // Charts are identified through the chart API; see readActiveShape for why they must not be read as shapes.
      const chart = context.workbook.getActiveChartOrNullObject();
      chart.load("isNullObject");
      await context.sync();
      if (!chart.isNullObject) {
        chart.load("name");
        await context.sync();
        return { label: chart.name, title: `Selected chart: ${chart.name}` };
      }
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
    return runExcel(async (context) => {
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
  return runExcel(async (context) => {
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

/** Sheets (and tables) described per round trip, so a workbook with hundreds of them never sends one giant request. */
const SYNC_GROUP = 20;

export async function getSheets(): Promise<unknown> {
  return runExcel(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/id,items/visibility");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    await context.sync();

    const usedRanges: Excel.Range[] = [];
    for (let start = 0; start < sheets.items.length; start += SYNC_GROUP) {
      const group = sheets.items.slice(start, start + SYNC_GROUP);
      const used = group.map((sheet) => sheet.getUsedRangeOrNullObject(true));
      used.forEach((range) => range.load("address"));
      group.forEach((sheet) => sheet.tables.load("items/name"));
      await context.sync();
      usedRanges.push(...used);
    }

    const tableRanges = sheets.items.map((sheet) =>
      sheet.tables.items.map((table) => {
        const range = table.getRange();
        return { name: table.name, range };
      }),
    );
    const pending = tableRanges.flat();
    for (let start = 0; start < pending.length; start += SYNC_GROUP) {
      pending.slice(start, start + SYNC_GROUP).forEach(({ range }) => range.load("address"));
      await context.sync();
    }

    return {
      activeSheet: active.name,
      sheets: sheets.items.map((sheet, index) => ({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility,
        usedRange: usedRanges[index]!.isNullObject ? null : stripSheet(usedRanges[index]!.address),
        isActive: sheet.name === active.name,
        tables: tableRanges[index]!.map(({ name, range }) => ({ name, range: stripSheet(range.address) })),
      })),
    };
  });
}

export async function readRange(input: ReadRangeInput): Promise<unknown> {
  return runExcel(async (context) => {
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
      if (input.includeText) page.load("text");
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
    const text = input.includeText ? (page ? (page.text as string[][]) : []) : undefined;
    const extra = text ? { text } : {};
    return input.format === "csv" ? { ...base, csv: toCsv(values), formulas, ...extra } : { ...base, values, formulas, ...extra };
  });
}

interface Match {
  sheet: string;
  address: string;
  value: CellValue;
  /** Present for formula cells. */
  formula?: string;
}

function buildMatcher(input: SearchDataInput): (text: string) => boolean {
  if (input.useRegex) {
    let pattern: RegExp;
    try {
      pattern = new RegExp(input.wholeCell ? `^(?:${input.searchTerm})$` : input.searchTerm, input.matchCase ? "" : "i");
    } catch (error) {
      throw new ToolFailure("invalid_arguments", `searchTerm is not a valid regular expression: ${error instanceof Error ? error.message : String(error)}`);
    }
    return (text) => pattern.test(text);
  }
  const needle = input.matchCase ? input.searchTerm : input.searchTerm.toLowerCase();
  return (text) => {
    const hay = input.matchCase ? text : text.toLowerCase();
    return input.wholeCell ? hay === needle : hay.includes(needle);
  };
}

/**
 * Above this many cells in the searched area, Excel's own Find does the work and only the hits travel to the pane,
 * instead of downloading every cell. Find matches the displayed text, so smaller areas keep the exact scan below.
 */
export const SEARCH_TUNING = { nativeCells: 100_000 };
/** Hit areas loaded per round trip. */
const FIND_GROUP = 20;

export async function searchData(input: SearchDataInput): Promise<unknown> {
  return runExcel(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();

    const targets = input.sheet
      ? [resolveSheet(context, input.sheet)]
      : sheets.items.filter((sheet) => sheet.visibility === "Visible");
    if (input.sheet) targets[0]!.load("name");

    const isMatch = buildMatcher(input);
    const wanted = input.offset + input.limit + 1; // one extra tells us whether more exist
    const matches: Match[] = [];
    let scanned = 0;
    let scanTruncated = false;
    let usedNativeFind = false;

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

      if (!input.address && !input.useRegex && !input.matchFormulas && used.rowCount * used.columnCount > SEARCH_TUNING.nativeCells) {
        // Excel's Find treats * ? ~ as wildcards; the search here is for the literal text.
        const term = input.searchTerm.replace(/[~*?]/g, "~$&");
        const found = sheet.findAllOrNullObject(term, { matchCase: input.matchCase, completeMatch: input.wholeCell });
        found.load("isNullObject,areas/items");
        await context.sync();
        usedNativeFind = true;
        if (found.isNullObject) continue;
        // No more areas than matches are needed; Excel does not promise their order, so put them in sheet order.
        const areas = found.areas.items.slice(0, wanted);
        for (let start = 0; start < areas.length; start += FIND_GROUP) {
          areas.slice(start, start + FIND_GROUP).forEach((area) => area.load("rowIndex,columnIndex,rowCount,columnCount"));
          await context.sync();
        }
        areas.sort((x, y) => x.rowIndex - y.rowIndex || x.columnIndex - y.columnIndex);
        for (let start = 0; start < areas.length; start += FIND_GROUP) {
          // One hit area can be a whole block of cells; only as many rows as could still be needed are read.
          const group = areas.slice(start, start + FIND_GROUP).map((area) => {
            const rows = Math.min(area.rowCount, Math.ceil(wanted / area.columnCount));
            return rows === area.rowCount ? area : area.getResizedRange(rows - area.rowCount, 0);
          });
          group.forEach((area) => area.load("values,formulas,rowIndex,columnIndex,rowCount,columnCount"));
          await context.sync();
          for (const area of group) {
            for (let r = 0; r < area.rowCount; r++) {
              for (let c = 0; c < area.columnCount; c++) {
                const value = area.values[r]?.[c];
                if (value === null || value === undefined || value === "") continue;
                const formula = String(area.formulas[r]?.[c] ?? "");
                matches.push({
                  sheet: sheet.name,
                  address: cellAddress(area.rowIndex + r, area.columnIndex + c),
                  value: value as CellValue,
                  ...(formula.startsWith("=") ? { formula } : {}),
                });
                if (matches.length >= wanted) break outer;
              }
            }
          }
        }
        continue;
      }

      // Read in row chunks so one call never loads an unbounded block.
      const rowsPerChunk = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / used.columnCount));
      for (let offset = 0; offset < used.rowCount; offset += rowsPerChunk) {
        if (scanned >= EXECUTOR_CAPS.searchScanCells) {
          scanTruncated = true;
          break outer;
        }
        const rows = Math.min(rowsPerChunk, used.rowCount - offset);
        const chunk = sheet.getRangeByIndexes(used.rowIndex + offset, used.columnIndex, rows, used.columnCount);
        chunk.load(input.matchFormulas ? "values,formulas" : "values");
        await context.sync();
        scanned += rows * used.columnCount;

        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < used.columnCount; c++) {
            const value = chunk.values[r]?.[c];
            const formula = input.matchFormulas ? String(chunk.formulas[r]?.[c] ?? "") : "";
            const isFormula = formula.startsWith("=");
            if (!isFormula && (value === null || value === undefined || value === "")) continue;
            if (isMatch(isFormula ? formula : String(value))) {
              matches.push({
                sheet: sheet.name,
                address: cellAddress(used.rowIndex + offset + r, used.columnIndex + c),
                value: value as CellValue,
                ...(isFormula ? { formula } : {}),
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
      ...(usedNativeFind ? { note: "Large sheet: searched with Excel's Find, which matches the text as displayed (for example 1,500, not 1500). Hits are listed in sheet order but only the first ones Excel returned are examined, so for a very common term narrow the search with address." } : {}),
    };
  });
}
