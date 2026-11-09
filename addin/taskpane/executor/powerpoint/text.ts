import type { EditSlideTextInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { TEXT_TYPES, tryEach } from "./model.js";
import { PPT_GEOMETRY, clip, normalizeText, parseShapeRef, requireSupport } from "./util.js";

interface Pending {
  ref: string;
  text: string;
  shape?: PowerPoint.Shape;
  name?: string;
  before?: string;
  error?: string;
}

/**
 * Replace the text of one or several shapes with `textFrame.textRange.text` (PowerPointApi 1.4). All shapes are checked
 * first and written in one sync; a shape that cannot take text is reported, the others are still edited.
 * Plain assignment can drop run-level formatting (mixed fonts, bold words); edit_slide_xml keeps it.
 */
export async function editSlideText(input: EditSlideTextInput): Promise<unknown> {
  requireSupport("edit_slide_text", PPT_GEOMETRY, "Editing shape text needs PowerPointApi 1.4.");
  const edits: Array<{ ref: string; text: string }> = input.edits ?? [{ ref: input.ref!, text: input.text! }];
  return PowerPoint.run(async (context) => {
    const items: Pending[] = edits.map((edit) => ({ ref: edit.ref, text: edit.text }));
    for (const item of items) {
      try {
        const { slideId, shapeId } = parseShapeRef(item.ref);
        const shape = context.presentation.slides.getItemOrNullObject(slideId).shapes.getItemOrNullObject(shapeId);
        shape.load("isNullObject,name,type");
        item.shape = shape;
      } catch (error) {
        item.error = error instanceof Error ? error.message : String(error);
      }
    }
    const lookups = items.filter((item) => item.shape);
    let lookedUp = true;
    try {
      await context.sync();
    } catch {
      lookedUp = false;
    }
    if (!lookedUp) {
      // A missing slide makes the whole batch fail; look the shapes up one by one to say which.
      for (const item of lookups) {
        try {
          const { slideId, shapeId } = parseShapeRef(item.ref);
          const shape = context.presentation.slides.getItemOrNullObject(slideId).shapes.getItemOrNullObject(shapeId);
          shape.load("isNullObject,name,type");
          await context.sync();
          item.shape = shape;
        } catch {
          item.error = `slide of ${item.ref} does not exist; read the current slide ids with get_presentation_metadata`;
          item.shape = undefined;
        }
      }
    }
    for (const item of items) {
      if (!item.shape) continue;
      if (item.shape.isNullObject) {
        item.error = `shape ${item.ref} does not exist; list the shapes with list_slide_shapes`;
        item.shape = undefined;
        continue;
      }
      item.name = item.shape.name;
      const type = String(item.shape.type);
      if (!TEXT_TYPES.has(type)) {
        item.error =
          type === "Table"
            ? `${item.ref} is a table; use edit_slide_table to change its cells`
            : `${item.ref} is a ${type} shape and has no text frame`;
        item.shape = undefined;
      }
    }

    const writable = items.filter((item) => item.shape);
    const loaded = await tryEach(context, writable, (item) => item.shape!.textFrame.textRange.load("text"));
    writable.forEach((item, i) => {
      if (loaded[i]) item.before = clip(normalizeText(item.shape!.textFrame.textRange.text), 200).text;
      else {
        item.error = `${item.ref} has no editable text frame (it may be a placeholder holding a picture, chart or table)`;
        item.shape = undefined;
      }
    });

    const ready = writable.filter((item) => item.shape);
    if (ready.length > 0) {
      try {
        for (const item of ready) item.shape!.textFrame.textRange.text = item.text;
        await context.sync();
      } catch {
        // Not atomic: some writes may have landed. Writing the same text again is harmless, so redo them one by one.
        for (const item of ready) {
          try {
            item.shape!.textFrame.textRange.text = item.text;
            await context.sync();
          } catch (error) {
            item.error = `could not write the text: ${error instanceof Error ? error.message : String(error)}`;
          }
        }
      }
    }

    const results = items.map((item) => (item.error ? { ref: item.ref, ok: false, error: item.error } : { ref: item.ref, ok: true, name: item.name, before: item.before, characters: item.text.length }));
    const edited = results.filter((r) => r.ok).length;
    if (edited === 0) {
      throw new ToolFailure("invalid_arguments", results.length === 1 ? (results[0] as { error: string }).error : `no shape was edited: ${(results[0] as { error: string }).error}`, { results });
    }
    return {
      summary: edited === 1 ? `Edited the text of ${results.find((r) => r.ok)!.ref}` : `Edited the text of ${edited} shapes${edited < results.length ? ` (${results.length - edited} skipped)` : ""}`,
      edited,
      failed: results.length - edited,
      results,
      note: "The text was replaced as plain text; formatting that differed inside the shape (mixed fonts, bold words, links) can be lost. Use edit_slide_xml for format-preserving edits.",
    };
  });
}
