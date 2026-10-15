import {
  columnToNumber,
  numberToColumn,
  type InsertRowsInput,
  type ModifyStructureInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet } from "./common.js";

export async function insertRows(input: InsertRowsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const rowNumber = Number.parseInt(input.at.replace(/[^0-9]/g, ""), 10);
    if (!Number.isFinite(rowNumber)) {
      throw new ToolFailure("invalid_arguments", `cannot parse row reference '${input.at}'`);
    }
    const start = input.position === "above" ? rowNumber : rowNumber + 1;
    const end = start + input.count - 1;
    sheet.getRange(`${start}:${end}`).insert(Excel.InsertShiftDirection.down);
    await context.sync();
    return { sheet: sheet.name, at: input.at, count: input.count, shiftedRange: `${start}:${end}` };
  });
}

export async function modifyStructure(input: ModifyStructureInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const isRows = input.dimension === "rows";
    let target: string;

    if (input.operation === "show_levels") {
      // Outline level to show: 1 collapses every group, higher levels expand progressively.
      if (isRows) sheet.showOutlineLevels(input.count, 0);
      else sheet.showOutlineLevels(0, input.count);
      target = `${input.dimension} outline level ${input.count}`;
    } else if (input.operation === "unfreeze") {
      sheet.freezePanes.unfreeze();
      target = "all panes";
    } else if (input.operation === "freeze") {
      if (isRows) sheet.freezePanes.freezeRows(input.count);
      else sheet.freezePanes.freezeColumns(input.count);
      target = `first ${input.count} ${input.dimension}`;
    } else {
      const ref = input.reference!;
      const first = isRows ? Number(ref) : columnToNumber(ref);
      // `insert` after the reference starts one row/column later.
      const start = input.operation === "insert" && input.position === "after" ? first + 1 : first;
      const end = start + input.count - 1;
      const address = isRows ? `${start}:${end}` : `${numberToColumn(start)}:${numberToColumn(end)}`;
      const range = sheet.getRange(address);
      switch (input.operation) {
        case "insert":
          range.insert(isRows ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
          break;
        case "delete":
          range.delete(isRows ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
          break;
        case "group":
          range.group(isRows ? Excel.GroupOption.byRows : Excel.GroupOption.byColumns);
          break;
        case "ungroup":
          range.ungroup(isRows ? Excel.GroupOption.byRows : Excel.GroupOption.byColumns);
          break;
        case "hide":
        case "unhide": {
          const hidden = input.operation === "hide";
          if (isRows) range.rowHidden = hidden;
          else range.columnHidden = hidden;
          break;
        }
      }
      target = address;
    }

    await context.sync();
    return { sheet: sheet.name, operation: input.operation, target };
  });
}
