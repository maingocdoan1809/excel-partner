import { isExcelToolName, type CellValue, type ToolResult } from "../../protocol/src/index.js";

/**
 * In-memory stand-in for the task pane's Office.js executor, used to drive
 * real provider CLIs end to end without Excel. Implements the seven tools on
 * a single dictionary-backed sheet model (enough to verify AI behaviour).
 */
export class FakeWorkbook {
  readonly sheets = new Map<string, Map<string, CellValue>>();
  readonly formulas = new Map<string, string>();
  readonly shapes = new Map<string, any>();
  readonly formats: Array<{ sheet: string; address: string; format: unknown }> = [];
  readonly calls: Array<{ tool: string; args: unknown; ok: boolean }> = [];
  selection = { sheet: "Sheet1", address: "A2:C2" };

  constructor() {
    const sheet = new Map<string, CellValue>();
    const rows: CellValue[][] = [
      ["Item", "Qty", "Price"],
      ["Apple", 3, 1.5],
      ["Pear", 2, 2],
    ];
    rows.forEach((row, r) => row.forEach((v, c) => sheet.set(`${colName(c)}${r + 1}`, v)));
    this.sheets.set("Sheet1", sheet);
  }

  get(sheet: string, address: string): CellValue {
    return this.sheets.get(sheet)?.get(address) ?? null;
  }

  async handle(tool: string, args: any): Promise<ToolResult> {
    const result = this.run(tool, args ?? {});
    this.calls.push({ tool, args, ok: result.ok });
    return result;
  }

  private run(tool: string, args: any): ToolResult {
    const sheetName: string = args.sheet?.name ?? this.selection.sheet;
    const sheet = this.sheets.get(sheetName);
    switch (tool) {
      case "get_selection": {
        const cells = expand(this.selection.address);
        return ok({
          workbookName: "fake.xlsx",
          sheet: this.selection.sheet,
          sheetId: "1",
          address: this.selection.address,
          rowCount: cells.length,
          columnCount: cells[0]?.length ?? 0,
          values: cells.map((row) => row.map((a) => this.get(this.selection.sheet, a))),
          headers: ["Item", "Qty", "Price"],
          truncated: false,
        });
      }
      case "get_sheets":
        return ok({
          sheets: [...this.sheets.keys()].map((name, i) => ({
            id: String(i + 1),
            name,
            visibility: "Visible",
            usedRange: "A1:C3",
            isActive: name === this.selection.sheet,
          })),
          activeSheet: this.selection.sheet,
        });
      case "read_range": {
        if (!sheet) return err("sheet_not_found", `sheet '${sheetName}' not found`);
        const cells = expand(args.address);
        return ok({
          sheet: sheetName,
          address: args.address,
          rowCount: cells.length,
          columnCount: cells[0]?.length ?? 0,
          values: cells.map((row) => row.map((a) => this.get(sheetName, a))),
          formulas: args.includeFormulas
            ? cells.map((row) => row.map((a) => this.formulas.get(`${sheetName}!${a}`) ?? String(this.get(sheetName, a) ?? "")))
            : null,
          truncated: false,
        });
      }
      case "write_range": {
        if (!sheet) return err("sheet_not_found", `sheet '${sheetName}' not found`);
        const { col, row } = parseCell(String(args.address).split(":")[0]!);
        (args.values as CellValue[][]).forEach((vals, r) =>
          vals.forEach((v, c) => sheet.set(`${colName(col + c)}${row + r}`, v)),
        );
        return ok({
          sheet: sheetName,
          address: args.address,
          rowsWritten: args.values.length,
          columnsWritten: args.values[0]?.length ?? 0,
        });
      }
      case "set_formula": {
        if (!sheet) return err("sheet_not_found", `sheet '${sheetName}' not found`);
        for (const cell of expand(args.address).flat()) {
          this.formulas.set(`${sheetName}!${cell}`, args.formula);
          sheet.set(cell, `=${String(args.formula).replace(/^=/, "")}`);
        }
        return ok({ sheet: sheetName, address: args.address, formula: args.formula });
      }
      case "format_range":
        this.formats.push({ sheet: sheetName, address: args.address, format: args.format });
        return ok({ sheet: sheetName, address: args.address, applied: Object.keys(args.format ?? {}) });
      case "insert_rows": {
        if (!sheet) return err("sheet_not_found", `sheet '${sheetName}' not found`);
        const at = Number(String(args.at).replace(/\D/g, ""));
        const start = args.position === "below" ? at + 1 : at;
        const moved = new Map<string, CellValue>();
        for (const [addr, v] of sheet) {
          const { col, row } = parseCell(addr);
          moved.set(`${colName(col)}${row >= start ? row + args.count : row}`, v);
        }
        sheet.clear();
        for (const [k, v] of moved) sheet.set(k, v);
        return ok({ sheet: sheetName, at: String(args.at), count: args.count, shiftedRange: `${start}:${start + args.count - 1}` });
      }
      case "add_shape":
        this.shapes.set(args.name ?? `shape${this.shapes.size + 1}`, args);
        return ok({ name: args.name, type: args.type, text: args.text ?? null, left: 0, top: 0, width: 1, height: 1 });
      case "delete_shapes":
        for (const name of args.names) this.shapes.delete(name);
        return ok({ sheet: sheetName, deleted: args.names, notFound: [] });
      case "list_shapes":
        return ok({ sheet: sheetName, count: this.shapes.size, shapes: [...this.shapes.keys()].map((name) => ({ name })) });
      case "add_connector":
      case "update_shape":
      case "clear_range":
      case "modify_structure":
      case "search_data":
        return ok({ sheet: sheetName, tool });
      default:
        // Advanced tools are covered by the Office.js mock tests; here they just acknowledge the call.
        return isExcelToolName(tool) ? ok({ sheet: sheetName, tool }) : err("unknown_tool", `unknown tool ${tool}`);
    }
  }
}

function ok(data: unknown): ToolResult {
  return { ok: true, data } as ToolResult;
}
function err(code: string, message: string): ToolResult {
  return { ok: false, error: { code, message } } as ToolResult;
}

export function colName(index: number): string {
  let n = index;
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function parseCell(a1: string): { col: number; row: number } {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(a1);
  if (!m) throw new Error(`bad address ${a1}`);
  let col = 0;
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) };
}

function expand(address: string): string[][] {
  const [a, b] = address.split(":");
  const start = parseCell(a!);
  const end = b ? parseCell(b) : start;
  const rows: string[][] = [];
  for (let r = start.row; r <= end.row; r++) {
    const row: string[] = [];
    for (let c = start.col; c <= end.col; c++) row.push(`${colName(c)}${r}`);
    rows.push(row);
  }
  return rows;
}
