import type { EditTableInput, InsertTableInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported } from "../../host/requirements.js";
import { assertParagraphIndex, clip, fitResult, requireSupport, WORD_API_1_3, wordRun } from "./common.js";

const CELL_READ_LIMIT = 500;

function htmlEscape(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function tableHtml(values: string[][], header: boolean): string {
  const rows = values.map((row, r) => `<tr>${row.map((cell) => `<${header && r === 0 ? "th" : "td"}>${htmlEscape(cell)}</${header && r === 0 ? "th" : "td"}>`).join("")}</tr>`);
  return `<table border="1" cellspacing="0" cellpadding="4">${rows.join("")}</table>`;
}

export async function insertTable(input: InsertTableInput): Promise<unknown> {
  const values = input.values;
  const rows = values.length;
  const columns = values[0]!.length;
  const header = input.header_row ?? true;
  return wordRun(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    const tables = isSupported(WORD_API_1_3) ? body.tables : undefined;
    tables?.load("items");
    await context.sync();
    const after = input.after_paragraph;
    if (after !== undefined) assertParagraphIndex(after, paragraphs.items.length, "after_paragraph");
    const tablesBefore = tables?.items.length;

    if (!tables) {
      // WordApi 1.1/1.2 has no Table API: insert an HTML table.
      const html = tableHtml(values, header);
      if (after !== undefined) paragraphs.items[after]!.insertParagraph("", "After").insertHtml(html, "Replace");
      else body.insertHtml(html, "End");
      await context.sync();
      return { summary: `Inserted a ${rows}×${columns} table (as HTML)`, rows, columns, table_index: null, method: "html", ...(input.style ? { note: "Table style was not applied: this Word has no Table API (WordApi 1.3)." } : {}) };
    }

    const table = after !== undefined
      ? paragraphs.items[after]!.insertTable(rows, columns, "After", values)
      : body.insertTable(rows, columns, "End", values);
    table.headerRowCount = header ? 1 : 0;
    await context.sync();

    let styleNote: string | undefined;
    if (input.style) {
      try {
        table.style = input.style;
        await context.sync();
      } catch {
        styleNote = `The table was inserted, but the style “${input.style}” could not be applied (check the exact style name).`;
      }
    }

    // Tables carry no id: at the end it is the last one; elsewhere match it by size and first cell.
    let tableIndex: number | null = null;
    if (after === undefined && tablesBefore !== undefined) tableIndex = tablesBefore;
    else {
      const now = body.tables;
      now.load("items/rowCount");
      await context.sync();
      const sized = now.items.map((t, i) => ({ t, i })).filter((x) => x.t.rowCount === rows);
      const firsts = sized.map((x) => { const cell = x.t.getCell(0, 0); cell.load("value"); return cell; });
      if (firsts.length > 0) await context.sync();
      const same = sized.filter((_, k) => firsts[k]!.value === values[0]![0]);
      if (same.length === 1) tableIndex = same[0]!.i;
    }
    return {
      summary: `Inserted a ${rows}×${columns} table${after !== undefined ? ` after paragraph ${after}` : " at the end"}`,
      rows,
      columns,
      table_index: tableIndex,
      method: "table_api",
      ...(tableIndex === null ? { note: "The new table's index could not be told apart from similar tables; read the document to find it." } : {}),
      ...(styleNote ? { style_note: styleNote } : {}),
    };
  });
}

function dims(values: string[][], rowCount: number): { rows: number; columns: number } {
  return { rows: rowCount, columns: values.reduce((max, row) => Math.max(max, row.length), 0) };
}

export async function editTable(input: EditTableInput): Promise<unknown> {
  requireSupport("Editing tables", WORD_API_1_3, "Edit the table in Word, or use execute_office_js.");
  if (input.operation === "merge") requireSupport("Merging table cells", ["WordApi", "1.4"]);
  return wordRun(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    if (input.table_index >= tables.items.length) {
      throw new ToolFailure("range_not_found", `table_index ${input.table_index} does not exist: the document has ${tables.items.length} table${tables.items.length === 1 ? "" : "s"} (0..${Math.max(0, tables.items.length - 1)}).`);
    }
    const table = tables.items[input.table_index]!;
    table.load("rowCount,values,headerRowCount,style");
    await context.sync();
    const { rows, columns } = dims(table.values, table.rowCount);
    const count = input.count ?? 1;
    const checkRow = (row: number | undefined, label = "row") => {
      if (row === undefined || row < 0 || row >= rows) throw new ToolFailure("invalid_arguments", `${label} ${row} is outside the table, which has ${rows} rows (0..${rows - 1}).`);
    };
    const checkColumn = (column: number | undefined, label = "column") => {
      if (column === undefined || column < 0 || column >= columns) throw new ToolFailure("invalid_arguments", `${label} ${column} is outside the table, which has ${columns} columns (0..${columns - 1}).`);
    };

    if (input.operation === "read") {
      const shown = table.values.map((r) => r.map((c) => clip(c, CELL_READ_LIMIT).text));
      return fitResult(shown.length, (n) => ({
        table_index: input.table_index,
        rows,
        columns,
        header_rows: table.headerRowCount,
        style: table.style,
        values: shown.slice(0, n),
        truncated: n < shown.length,
      }));
    }

    let summary: string;
    switch (input.operation) {
      case "set_cell":
        checkRow(input.row);
        checkColumn(input.column);
        table.getCell(input.row!, input.column!).value = input.text!;
        summary = `Set cell (${input.row}, ${input.column}) of table ${input.table_index}`;
        break;
      case "add_rows":
        checkRow(input.row);
        const tableRows = table.rows;
        tableRows.load("items");
        await context.sync();
        tableRows.items[input.row!]!.insertRows(input.position === "before" ? "Before" : "After", count);
        summary = `Added ${count} row${count === 1 ? "" : "s"} ${input.position ?? "after"} row ${input.row}`;
        break;
      case "delete_rows":
        checkRow(input.row);
        if (input.row! + count > rows) throw new ToolFailure("invalid_arguments", `Cannot delete ${count} rows from row ${input.row}: the table has ${rows} rows.`);
        if (count >= rows) throw new ToolFailure("invalid_arguments", "That would delete every row; use operation delete_table to remove the table.");
        table.deleteRows(input.row!, count);
        summary = `Deleted ${count} row${count === 1 ? "" : "s"} from row ${input.row}`;
        break;
      case "add_columns":
        checkColumn(input.column);
        table.getCell(0, input.column!).insertColumns(input.position === "before" ? "Before" : "After", count);
        summary = `Added ${count} column${count === 1 ? "" : "s"} ${input.position ?? "after"} column ${input.column}`;
        break;
      case "delete_columns":
        checkColumn(input.column);
        if (input.column! + count > columns) throw new ToolFailure("invalid_arguments", `Cannot delete ${count} columns from column ${input.column}: the table has ${columns} columns.`);
        if (count >= columns) throw new ToolFailure("invalid_arguments", "That would delete every column; use operation delete_table to remove the table.");
        table.deleteColumns(input.column!, count);
        summary = `Deleted ${count} column${count === 1 ? "" : "s"} from column ${input.column}`;
        break;
      case "merge":
        checkRow(input.row);
        checkRow(input.to_row, "to_row");
        checkColumn(input.column);
        checkColumn(input.to_column, "to_column");
        if (input.to_row! < input.row! || input.to_column! < input.column!) throw new ToolFailure("invalid_arguments", "to_row / to_column must not be before row / column.");
        table.mergeCells(input.row!, input.column!, input.to_row!, input.to_column!);
        summary = `Merged cells (${input.row}, ${input.column}) to (${input.to_row}, ${input.to_column})`;
        break;
      case "delete_table":
        table.delete();
        await context.sync();
        return { summary: `Deleted table ${input.table_index}`, deleted: true, table_index: input.table_index, note: "Table indexes after this one have shifted." };
      case "set_style":
        table.style = input.style!;
        summary = `Set table ${input.table_index} style to “${input.style}”`;
        break;
      default:
        throw new ToolFailure("invalid_arguments", `Unknown table operation ${String(input.operation)}.`);
    }
    await context.sync();
    table.load("rowCount,values");
    await context.sync();
    const after = dims(table.values, table.rowCount);
    return { summary, table_index: input.table_index, rows: after.rows, columns: after.columns };
  });
}
