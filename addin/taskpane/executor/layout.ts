import type { ManageNamesInput, MergeRangeInput, PageLayoutInput } from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";
import { findOccupied, loadGeometry, overwriteFailure } from "./write.js";

export async function mergeRange(input: MergeRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    const geometry = await loadGeometry(context, range);

    if (input.operation === "merge") {
      if (geometry.rowCount * geometry.columnCount < 2) throw new ToolFailure("invalid_arguments", "merge needs at least two cells");
      if (!input.allowOverwrite) {
        // Excel keeps only the top-left value (per row when merging across); everything else is lost.
        const keepers = (r: number, c: number) =>
          input.across ? c === geometry.columnIndex : r === geometry.rowIndex && c === geometry.columnIndex;
        const found = await findOccupied(context, [{ range, geometry }], keepers);
        if (found.total > 0) throw overwriteFailure(found);
      }
      range.merge(input.across);
    } else {
      range.unmerge();
    }
    range.load("address");
    await context.sync();
    return { sheet: sheet.name, address: stripSheet(range.address), operation: input.operation };
  });
}

export async function manageNames(input: ManageNamesInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheetScoped = input.scope !== "workbook";
    const sheet = sheetScoped ? context.workbook.worksheets.getItem(input.scope) : null;
    const names = sheet ? sheet.names : context.workbook.names;

    if (input.operation === "list") {
      const all = [{ scope: "workbook", names: context.workbook.names }];
      context.workbook.worksheets.load("items/name");
      await context.sync();
      for (const s of context.workbook.worksheets.items) all.push({ scope: s.name, names: s.names });
      all.forEach((entry) => entry.names.load("items/name,items/formula,items/visible,items/comment"));
      await context.sync();
      return {
        names: all.flatMap((entry) =>
          entry.names.items.filter((n) => n.visible).map((n) => ({ name: n.name, reference: n.formula, scope: entry.scope, comment: n.comment || null })),
        ),
      };
    }

    const existing = names.getItemOrNullObject(input.name!);
    existing.load("isNullObject");
    await context.sync();

    if (input.operation === "delete") {
      if (existing.isNullObject) throw new ToolFailure("invalid_arguments", `no name '${input.name}' at ${input.scope} scope`);
      existing.delete();
      await context.sync();
      return { operation: "delete", name: input.name, scope: input.scope };
    }

    if (!existing.isNullObject) {
      throw new ToolFailure("invalid_arguments", `name '${input.name}' already exists at ${input.scope} scope; delete it first`);
    }
    const reference = input.reference!.startsWith("=") ? input.reference! : `=${input.reference!}`;
    names.add(input.name!, reference, input.comment);
    await context.sync();
    return { operation: "add", name: input.name, reference, scope: input.scope };
  });
}

const PAPER = { Letter: "Letter", Legal: "Legal", A3: "A3", A4: "A4", A5: "A5", Tabloid: "Tabloid" } as const;

export async function pageLayout(input: PageLayoutInput): Promise<unknown> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    throw new ToolFailure("unsupported", "this Excel version cannot change page layout (needs ExcelApi 1.9)");
  }
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const layout = sheet.pageLayout;
    const applied: string[] = [];
    const did = (name: string, run: () => void) => {
      run();
      applied.push(name);
    };

    if (input.printArea) did("printArea", () => layout.setPrintArea(sheet.getRange(input.printArea!)));
    if (input.orientation) did("orientation", () => (layout.orientation = input.orientation === "landscape" ? Excel.PageOrientation.landscape : Excel.PageOrientation.portrait));
    if (input.paperSize) did("paperSize", () => (layout.paperSize = PAPER[input.paperSize!] as Excel.PaperType));
    if (input.scalePercent !== undefined) did("scalePercent", () => (layout.zoom = { scale: input.scalePercent }));
    else if (input.fitToPagesWide !== undefined || input.fitToPagesTall !== undefined) {
      did("fitToPages", () => (layout.zoom = { horizontalFitToPages: input.fitToPagesWide ?? 1, verticalFitToPages: input.fitToPagesTall || undefined }));
    }
    if (input.repeatTitleRows) did("repeatTitleRows", () => layout.setPrintTitleRows(sheet.getRange(input.repeatTitleRows!)));
    if (input.repeatTitleColumns) did("repeatTitleColumns", () => layout.setPrintTitleColumns(sheet.getRange(input.repeatTitleColumns!)));
    const m = input.marginsInches;
    if (m) {
      did("margins", () => {
        if (m.left !== undefined) layout.leftMargin = m.left * 72;
        if (m.right !== undefined) layout.rightMargin = m.right * 72;
        if (m.top !== undefined) layout.topMargin = m.top * 72;
        if (m.bottom !== undefined) layout.bottomMargin = m.bottom * 72;
      });
    }
    const pages = layout.headersFooters.defaultForAllPages;
    if (input.header) {
      did("header", () => {
        if (input.header!.left !== undefined) pages.leftHeader = input.header!.left;
        if (input.header!.center !== undefined) pages.centerHeader = input.header!.center;
        if (input.header!.right !== undefined) pages.rightHeader = input.header!.right;
      });
    }
    if (input.footer) {
      did("footer", () => {
        if (input.footer!.left !== undefined) pages.leftFooter = input.footer!.left;
        if (input.footer!.center !== undefined) pages.centerFooter = input.footer!.center;
        if (input.footer!.right !== undefined) pages.rightFooter = input.footer!.right;
      });
    }
    if (input.centerHorizontally !== undefined) did("centerHorizontally", () => (layout.centerHorizontally = input.centerHorizontally!));
    if (input.centerVertically !== undefined) did("centerVertically", () => (layout.centerVertically = input.centerVertically!));
    if (input.printGridlines !== undefined) did("printGridlines", () => (layout.printGridlines = input.printGridlines!));
    if (input.printHeadings !== undefined) did("printHeadings", () => (layout.printHeadings = input.printHeadings!));

    await context.sync();
    return { sheet: sheet.name, applied };
  });
}

