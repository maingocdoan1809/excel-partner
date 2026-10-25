import { MAX_REPLACE_CELLS, type FindReplaceInput } from "../../../protocol/src/index.js";
import { cellAddress, scanCells } from "./audit.js";
import { ToolFailure, isErrorValue } from "./common.js";

interface Change {
  sheet: string;
  address: string;
  before: string;
  after: string;
  isFormula: boolean;
}

const escapeRegExp = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Text Excel would otherwise turn into a formula, number or date when written back as a value. */
function needsTextPrefix(text: string): boolean {
  if (text.startsWith("=")) return true;
  const trimmed = text.trim();
  if (trimmed === "") return false;
  return !Number.isNaN(Number(trimmed)) || /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/.test(trimmed);
}

export async function findReplace(input: FindReplaceInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const source = escapeRegExp(input.find);
    const flags = input.matchCase ? "g" : "gi";
    const pattern = new RegExp(input.wholeCell ? `^${source}$` : source, flags);
    const replaceWith = () => input.replace;

    const changes: Change[] = [];
    let total = 0;

    const { scanTruncated } = await scanCells(context, input, (chunk) => {
      for (let r = 0; r < chunk.rows; r++) {
        for (let c = 0; c < chunk.columns; c++) {
          const raw = String(chunk.formulas[r]?.[c] ?? "");
          const isFormula = raw.startsWith("=");
          let before: string;
          if (isFormula) {
            if (!input.inFormulas) continue;
            before = raw;
          } else {
            const value = chunk.values[r]?.[c];
            if (typeof value !== "string" || isErrorValue(value)) continue;
            before = value;
          }
          pattern.lastIndex = 0;
          if (!pattern.test(before)) continue;
          pattern.lastIndex = 0;
          const after = before.replace(pattern, replaceWith);
          if (after === before) continue;
          total++;
          if (changes.length <= MAX_REPLACE_CELLS) {
            changes.push({ sheet: chunk.sheet, address: cellAddress(chunk.rowIndex + r, chunk.columnIndex + c), before, after, isFormula });
          }
        }
      }
    });

    if (!input.apply) {
      return {
        applied: false,
        totalMatches: total,
        returned: Math.min(total, input.limit, changes.length),
        hasMore: total > input.limit,
        changes: changes.slice(0, input.limit).map(({ sheet, address, before, after }) => ({ sheet, address, before, after })),
        scanTruncated,
        message:
          total === 0
            ? "Nothing to replace."
            : `Preview only; nothing was changed. Call again with apply=true to replace in ${total} cell(s).${total > MAX_REPLACE_CELLS ? ` That is more than the ${MAX_REPLACE_CELLS}-cell limit per call: narrow the range first.` : ""}`,
      };
    }

    if (total > MAX_REPLACE_CELLS) {
      throw new ToolFailure(
        "size_limit_exceeded",
        `${total} cells would change (limit ${MAX_REPLACE_CELLS} per call). Narrow the range with sheet/address and retry.`,
      );
    }
    if (scanTruncated) {
      throw new ToolFailure("size_limit_exceeded", "The workbook is too large to scan in one call. Narrow the range with sheet/address and retry.");
    }

    const bySheet = new Map<string, Change[]>();
    for (const change of changes) bySheet.set(change.sheet, [...(bySheet.get(change.sheet) ?? []), change]);

    const formulaCells: Array<{ sheet: string; address: string; range: Excel.Range }> = [];
    for (const [name, list] of bySheet) {
      const sheet = context.workbook.worksheets.getItem(name);
      for (const change of list) {
        const range = sheet.getRange(change.address);
        if (change.isFormula) {
          range.formulas = [[change.after]];
          if (formulaCells.length < 200) formulaCells.push({ sheet: name, address: change.address, range });
        } else {
          // A replaced text that starts with "=" must stay text, not turn into a formula.
          range.values = [[needsTextPrefix(change.after) ? `'${change.after}` : change.after]];
        }
      }
    }
    formulaCells.forEach((cell) => cell.range.load("values"));
    await context.sync();

    const errors = formulaCells
      .filter((cell) => isErrorValue(cell.range.values[0]?.[0]))
      .map((cell) => `${cell.sheet}!${cell.address}=${cell.range.values[0]![0]}`);

    return {
      applied: true,
      replaced: changes.length,
      sheets: [...bySheet.keys()],
      changes: changes.slice(0, 20).map(({ sheet, address, before, after }) => ({ sheet, address, before, after })),
      ...(errors.length > 0 ? { messages: [`${errors.length} edited formula(s) now evaluate to an error: ${errors.slice(0, 5).join(", ")}. Fix or revert them.`] } : {}),
    };
  });
}
