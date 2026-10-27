import type { InsertSheetsFromFileInput, ManageSheetInput } from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";

async function assertNameFree(context: Excel.RequestContext, name: string): Promise<void> {
  const clash = context.workbook.worksheets.getItemOrNullObject(name);
  clash.load("isNullObject");
  await context.sync();
  if (!clash.isNullObject) throw new ToolFailure("invalid_arguments", `a sheet named '${name}' already exists`);
}

export async function manageSheet(input: ManageSheetInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    if (input.operation === "add") {
      if (input.name) await assertNameFree(context, input.name);
      const added = input.name ? sheets.add(input.name) : sheets.add();
      if (input.position !== undefined) added.position = input.position;
      added.load("name,position");
      await context.sync();
      return { operation: "add", sheet: added.name, position: added.position };
    }

    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name,position,visibility");
    await context.sync();
    const result: Record<string, unknown> = { operation: input.operation, sheet: sheet.name };

    switch (input.operation) {
      case "delete": {
        if (!input.confirmDelete) {
          const used = sheet.getUsedRangeOrNullObject(true);
          used.load("isNullObject,address");
          await context.sync();
          const what = used.isNullObject ? "an empty sheet" : `its data (${stripSheet(used.address)})`;
          throw new ToolFailure(
            "confirmation_required",
            `Deleting sheet '${sheet.name}' removes ${what}. Ask the user to confirm, then retry with confirmDelete=true.`,
            { sheet: sheet.name, usedRange: used.isNullObject ? null : stripSheet(used.address) },
          );
        }
        sheets.load("items/visibility");
        await context.sync();
        if (sheets.items.filter((s) => s.visibility === "Visible").length <= 1 && sheet.visibility === "Visible") {
          throw new ToolFailure("invalid_arguments", "cannot delete the last visible sheet");
        }
        sheet.delete();
        break;
      }
      case "rename": {
        await assertNameFree(context, input.newName!);
        sheet.name = input.newName!;
        result.sheet = input.newName;
        result.previousName = sheet.name;
        break;
      }
      case "duplicate": {
        if (input.name) await assertNameFree(context, input.name);
        const copy = sheet.copy(Excel.WorksheetPositionType.after, sheet);
        if (input.name) copy.name = input.name;
        if (input.position !== undefined) copy.position = input.position;
        copy.load("name,position");
        await context.sync();
        result.duplicate = copy.name;
        result.position = copy.position;
        return result;
      }
      case "move":
        sheet.position = input.position!;
        break;
      case "activate":
        sheet.activate();
        break;
      case "hide":
      case "show": {
        if (input.operation === "hide") {
          sheets.load("items/visibility");
          await context.sync();
          if (sheets.items.filter((s) => s.visibility === "Visible").length <= 1) {
            throw new ToolFailure("invalid_arguments", "cannot hide the last visible sheet");
          }
        }
        sheet.visibility = input.operation === "hide" ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;
        break;
      }
      case "set_tab_color":
        sheet.tabColor = input.tabColor === "none" ? "" : input.tabColor!;
        break;
    }
    await context.sync();
    return result;
  });
}

/**
 * Copy sheets from a closed workbook into this one (formatting, merges and shapes come along). The bridge
 * reads the file and sends its bytes as `fileBase64`; the pane never sees file paths.
 */
export async function insertSheetsFromFile(input: InsertSheetsFromFileInput & { fileBase64?: string }): Promise<unknown> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
    throw new ToolFailure("unsupported", "this Excel version cannot insert sheets from a file (needs ExcelApi 1.13)");
  }
  if (!input.fileBase64) throw new ToolFailure("invalid_arguments", "the file contents did not reach the task pane");
  if (input.position === "after" && !input.afterSheet) {
    throw new ToolFailure("invalid_arguments", "position `after` needs afterSheet");
  }

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const before = new Set(sheets.items.map((s) => s.name));
    if (input.afterSheet && !before.has(input.afterSheet)) {
      throw new ToolFailure("sheet_not_found", `sheet '${input.afterSheet}' not found in the open workbook`, { available: [...before] });
    }

    const positionType = { end: "End", beginning: "Beginning", after: "After" }[input.position] as Excel.WorksheetPositionType;
    // Importing a whole workbook in one call has crashed Excel, so go one sheet at a time (the bridge lists every
    // sheet when the caller names none). Each later sheet goes after the previous one to keep the file's order.
    const names = input.sheets ?? [undefined];
    const inserted: string[] = [];
    let known = new Set(before);
    let anchor: string | undefined;
    for (const [index, name] of names.entries()) {
      const placeAfter = index > 0 && anchor !== undefined;
      try {
        context.workbook.insertWorksheetsFromBase64(input.fileBase64!, {
          positionType: placeAfter ? "After" as Excel.WorksheetPositionType : positionType,
          ...(placeAfter ? { relativeTo: anchor } : input.position === "after" ? { relativeTo: input.afterSheet } : {}),
          ...(name === undefined ? {} : { sheetNamesToInsert: [name] }),
        });
        await context.sync();
        sheets.load("items/name");
        await context.sync();
      } catch (err) {
        if (inserted.length === 0) throw err;
        const reason = err instanceof Error ? err.message : String(err);
        throw new ToolFailure("office_error", `inserted ${inserted.length} of ${names.length} sheets, then '${name}' failed: ${reason}`, { inserted, failedSheet: name });
      }
      const added = sheets.items.map((s) => s.name).filter((sheetName) => !known.has(sheetName));
      inserted.push(...added);
      known = new Set([...known, ...added]);
      anchor = added.at(-1) ?? anchor;
      // Let Excel repaint between sheets so the window does not freeze for the whole import.
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    return { inserted, count: inserted.length };
  });
}
