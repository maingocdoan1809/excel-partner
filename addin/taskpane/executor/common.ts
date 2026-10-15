import type { CellValue, SheetRef, ToolError } from "../../../protocol/src/index.js";

/** A failure the executor raises on purpose, with a code the model can act on. */
export class ToolFailure extends Error {
  constructor(
    readonly code: ToolError["code"],
    message: string,
    readonly detail?: unknown,
  ) {
    super(message);
    this.name = "ToolFailure";
  }
}

export function resolveSheet(context: Excel.RequestContext, ref?: SheetRef): Excel.Worksheet {
  if (ref?.id) return context.workbook.worksheets.getItem(ref.id);
  if (ref?.name) return context.workbook.worksheets.getItem(ref.name);
  return context.workbook.worksheets.getActiveWorksheet();
}

/** `Sheet1!A1:B2` / `'My Sheet'!A1` -> `A1:B2`. The sheet is reported separately. */
export function stripSheet(address: string): string {
  const bang = address.lastIndexOf("!");
  return bang === -1 ? address : address.slice(bang + 1);
}

export function normalizeValues(values: unknown[][] | undefined): CellValue[][] {
  if (!values) return [];
  return values.map((row) => row.map((cell) => (cell === undefined ? null : (cell as CellValue))));
}

/** Keep at most `limit` cells, cutting whole rows first. */
export function truncateCells<T>(rows: T[][], limit: number): T[][] {
  let remaining = limit;
  const out: T[][] = [];
  for (const row of rows) {
    if (remaining <= 0) break;
    const slice = row.slice(0, remaining);
    out.push(slice);
    remaining -= slice.length;
  }
  return out;
}

/** RFC 4180 CSV: quote fields containing `,` `"` or a line break; empty cells stay empty. */
export function toCsv(rows: CellValue[][]): string {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          if (cell === null || cell === undefined) return "";
          const text = String(cell);
          return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
        })
        .join(","),
    )
    .join("\n");
}

export function workbookName(): string {
  try {
    const url = Office.context.document.url;
    if (!url) return "unsaved";
    const parts = url.split(/[\\/]/);
    return decodeURIComponent(parts[parts.length - 1] ?? "") || "unsaved";
  } catch {
    return "unknown";
  }
}

export function toToolError(error: unknown): ToolError {
  if (error instanceof ToolFailure) {
    return error.detail === undefined
      ? { code: error.code, message: error.message }
      : { code: error.code, message: error.message, detail: error.detail };
  }
  const e = error as { code?: string; message?: string; debugInfo?: unknown };
  const message = e?.message ?? String(error);
  switch (e?.code) {
    case "ItemNotFound":
      return { code: "sheet_not_found", message, detail: { officeCode: e.code } };
    case "InvalidArgument":
    case "InvalidReference":
    case "BadRequest":
      return { code: "invalid_arguments", message, detail: { officeCode: e.code } };
    case "AccessDenied":
    case "ReadOnly":
      return { code: "workbook_readonly", message, detail: { officeCode: e.code } };
    case "ApiNotFound":
    case "NotImplemented":
      return { code: "unsupported", message, detail: { officeCode: e.code } };
    default:
      return typeof e?.code === "string"
        ? { code: "office_error", message, detail: { officeCode: e.code, debugInfo: e.debugInfo } }
        : { code: "office_error", message };
  }
}
