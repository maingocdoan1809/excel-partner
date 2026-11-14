import { sameWorkbookName } from "../../../protocol/src/index.js";
import { activeHost } from "../host/index.js";
import { showToast } from "./toast.js";

interface Place { sheet?: string; address: string }

const placeLabel = (target: Place & { file?: string }) =>
  `${target.file ? `${target.file} → ` : ""}${target.sheet ? `${target.sheet}!` : ""}${target.address}`;

/** Moving to a cell: in this workbook, or, for a link that names another open file, in that file's window. */
export function setupCellNavigation(deps: { readName(): Promise<string | null>; goToFile(target: Place & { file: string }): Promise<void> }) {
  /** Show a range in this workbook: its sheet (or the open one) is activated and the range selected. Throws when it cannot. */
  async function selectRange(target: Place): Promise<void> {
    if (!activeHost().hasSheets) throw new Error(`Links to cells only work in Excel; this is ${activeHost().app}.`);
    await Excel.run(async (context) => {
      const sheet = target.sheet ? context.workbook.worksheets.getItemOrNullObject(target.sheet) : context.workbook.worksheets.getActiveWorksheet();
      if (target.sheet) {
        await context.sync();
        if ((sheet as Excel.Worksheet & { isNullObject?: boolean }).isNullObject) throw new Error(`There is no sheet named "${target.sheet}" in this workbook.`);
      }
      sheet.activate();
      sheet.getRange(target.address).select();
      await context.sync();
    });
  }

  async function goTo(target: Place): Promise<void> {
    try {
      await selectRange(target);
    } catch (error) {
      showToast(`Cannot go to ${placeLabel(target)}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /** A link in a reply: in this workbook, or, when it names another open file, in that file's window (the bridge asks its pane). */
  async function goToLink(target: Place & { file?: string }): Promise<void> {
    if (!target.file) return goTo(target);
    const here = await deps.readName();
    if (here && sameWorkbookName(here, target.file)) return goTo({ ...(target.sheet ? { sheet: target.sheet } : {}), address: target.address });
    try {
      await deps.goToFile({ file: target.file, ...(target.sheet ? { sheet: target.sheet } : {}), address: target.address });
    } catch (error) {
      showToast(`Cannot go to ${placeLabel(target)}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  return { selectRange, goTo, goToLink };
}
