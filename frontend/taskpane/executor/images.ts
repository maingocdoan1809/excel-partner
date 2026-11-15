import type { ResolvedInsertImageInput } from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, runExcel } from "./common.js";
import { centreIn, fitImage, naturalSize, type Box, type Size } from "./image-layout.js";

/** Without a size, a picture is at most this big, so a photo does not cover the whole sheet. */
const DEFAULT_LIMIT: Size = { width: 480, height: 360 };

/** insert_image in Excel: a floating picture on a sheet. */
export async function insertImage(input: ResolvedInsertImageInput): Promise<unknown> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    throw new ToolFailure("unsupported", "this Excel version cannot insert pictures (needs ExcelApi 1.9)");
  }
  return runExcel(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");

    // Where it goes: a range to fill, explicit points, or the selected cell.
    const anchor = input.anchor ? sheet.getRange(input.anchor) : null;
    anchor?.load("left,top,width,height");
    const selection = !anchor && input.left === undefined && input.top === undefined && !input.sheet ? context.workbook.getSelectedRange() : null;
    selection?.load("left,top");
    const name = input.name ?? `Image_${Math.random().toString(36).slice(2, 7)}`;
    const existing = sheet.shapes.getItemOrNullObject(name);
    existing.load("isNullObject");
    await context.sync();
    if (!existing.isNullObject) throw new ToolFailure("invalid_arguments", `a shape named '${name}' already exists; pick another name`);

    const shape = sheet.shapes.addImage(input.imageBase64);
    shape.name = name;
    shape.load("width,height");
    await context.sync();

    // The size Excel gave it is the best measure of the picture's own size.
    const own: Size = shape.width > 0 && shape.height > 0 ? { width: shape.width, height: shape.height } : naturalSize(input) ?? { width: 0, height: 0 };
    const box: Box = anchor
      ? centreIn(anchor, fitImage(own, { width: anchor.width, height: anchor.height }))
      : { ...fitImage(own, { width: input.width, height: input.height }, DEFAULT_LIMIT), left: input.left ?? selection?.left ?? 0, top: input.top ?? selection?.top ?? 0 };
    shape.lockAspectRatio = false;
    shape.left = box.left;
    shape.top = box.top;
    shape.width = box.width;
    shape.height = box.height;
    shape.lockAspectRatio = true;
    await context.sync();

    let note: string | undefined;
    if (input.alt_text) {
      try {
        shape.altTextDescription = input.alt_text;
        await context.sync();
      } catch {
        note = "The picture was inserted, but this Excel cannot set alt text.";
      }
    }
    return {
      summary: `Inserted picture "${name}" on ${sheet.name} (${input.sourceLabel})`,
      sheet: sheet.name,
      name,
      left: box.left,
      top: box.top,
      width: box.width,
      height: box.height,
      source: input.sourceLabel,
      ...(note ? { note } : {}),
    };
  });
}
