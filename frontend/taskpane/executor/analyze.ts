import {
  AnalyzeError,
  EXECUTOR_CAPS,
  RangeAnalyzer,
  columnLetter,
  isDateFormat,
  type AnalyzeColumn,
  type AnalyzeRangeInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, runExcel, stripSheet } from "./common.js";

/**
 * `analyze_range`: read the table in blocks that fit one Office.js sync, hand each block to the analyzer and keep only its
 * running totals, so a 100,000-row sheet costs the conversation a few lines. The arithmetic is in protocol/src/analyze.ts.
 */
export async function analyzeRange(input: AnalyzeRangeInput): Promise<unknown> {
  return runExcel(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const full = sheet.getRange(input.address);
    full.load("address,rowCount,columnCount,rowIndex,columnIndex");
    await context.sync();

    const headerRows = input.headerRow ? 1 : 0;
    if (full.rowCount <= headerRows) throw new ToolFailure("invalid_arguments", `${stripSheet(full.address)} has no data rows${headerRows ? " below the header row" : ""}.`);
    const header = input.headerRow ? sheet.getRangeByIndexes(full.rowIndex, full.columnIndex, 1, full.columnCount) : null;
    // The number format of the first data row tells which columns hold dates; one row is all it takes.
    const formats = sheet.getRangeByIndexes(full.rowIndex + headerRows, full.columnIndex, 1, full.columnCount);
    header?.load("values");
    formats.load("numberFormat");
    await context.sync();

    const columns: AnalyzeColumn[] = Array.from({ length: full.columnCount }, (_, index) => {
      const letter = columnLetter(full.columnIndex + index);
      const title = header ? String(header.values[0]?.[index] ?? "").trim() : "";
      return { index, letter, name: title || letter, dateLike: isDateFormat(formats.numberFormat[0]?.[index]) };
    });

    let analyzer: RangeAnalyzer;
    try { analyzer = new RangeAnalyzer(columns, input); }
    catch (error) { throw error instanceof AnalyzeError ? new ToolFailure("invalid_arguments", error.message) : error; }

    const dataRows = full.rowCount - headerRows;
    const scanRows = Math.min(dataRows, Math.max(1, Math.floor(EXECUTOR_CAPS.searchScanCells / full.columnCount)));
    const blockRows = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / full.columnCount));
    try {
      for (let start = 0; start < scanRows; start += blockRows) {
        const rows = Math.min(blockRows, scanRows - start);
        const block = sheet.getRangeByIndexes(full.rowIndex + headerRows + start, full.columnIndex, rows, full.columnCount);
        block.load("values,valueTypes");
        await context.sync();
        analyzer.add(block.values, block.valueTypes, full.rowIndex + headerRows + start + 1);
      }
      const result = analyzer.result();
      const truncated = scanRows < dataRows;
      if (truncated) (result.notes as string[]).push(`Only the first ${scanRows} of ${dataRows} data rows were scanned (cell limit); narrow the address to cover the rest.`);
      return { sheet: sheet.name, address: stripSheet(full.address), mode: input.mode, headerRow: input.headerRow, dataRows, truncated, ...result };
    } catch (error) {
      throw error instanceof AnalyzeError ? new ToolFailure("invalid_arguments", error.message) : error;
    }
  });
}
