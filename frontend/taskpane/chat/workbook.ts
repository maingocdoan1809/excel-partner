import { activeHost } from "../host/index.js";
import { redrawMessages } from "./messages.js";
import { app } from "./state.js";

export async function currentWorkbookName(): Promise<string> {
  return (await activeHost().readName()) ?? "";
}

/**
 * Read the sheet names from Excel. True when they changed. With `redraw`, a changed list also draws the chat again, since which
 * cells of a reply are linked depends on it (a chat opened before the first read, for one).
 */
export async function refreshSheetNames(redraw = true): Promise<boolean> {
  if (!activeHost().hasSheets) return false;
  try {
    const names = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((sheet) => sheet.name);
    });
    if (names.join("\n") === app.sheetNames.join("\n")) return false;
    app.sheetNames = names;
    if (redraw && !app.currentRequestId && app.session.messages.length) redrawMessages();
    return true;
  } catch {
    return false; // no Excel behind this page, or it is busy: keep the names we have
  }
}

/** What Office calls the open file now (it changes on Save As), or null when it cannot be read. There is no event for a rename. */
export async function readWorkbookName(): Promise<string | null> {
  return activeHost().readName();
}

/**
 * Tell the bridge the name once on connect, so the file is listed with its name at once. After that the bridge asks
 * (`readWorkbookName`) when someone looks at the list, so nothing here runs on a timer.
 */
export async function publishWorkbookName(): Promise<void> {
  const name = await readWorkbookName();
  if (name && app.client?.connected) app.client.sendWorkbookInfo(name);
}
