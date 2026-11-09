import { PptToolInputSchemas, type PptToolName, type ToolResult } from "../../../../protocol/src/index.js";
import { ToolFailure, toToolError } from "../common.js";
import { editSlideText } from "./text.js";
import { getPresentationMetadata, listLayouts, listSlideShapes, readSlideText } from "./read.js";
import { alignShapes, groupShapes, insertSlideElement, removeSlideElement, setShapeAltText, setShapeZOrder, setSlideHyperlink, ungroupShapes } from "./shapes.js";
import { addSlide, duplicateSlide, modifyPresentationStructure, moveSlide, setSlideBackground } from "./slides.js";
import { editSlideTable, insertSlideTable } from "./tables.js";
import { capResult, requirePowerPoint, resultSize, MAX_RESULT_CHARS } from "./util.js";
import { exportSlideImage, verifySlideVisual, verifySlides } from "./verify.js";
import { editSlideXml } from "./xml.js";

export { getPptSelection, getPptSelectionLabel, readPptPresentationName } from "./selection.js";

class UnknownTool extends Error {}

const HANDLERS: Record<PptToolName, (args: never) => Promise<unknown>> = {
  get_presentation_metadata: () => getPresentationMetadata(),
  list_layouts: () => listLayouts(),
  add_slide: addSlide,
  list_slide_shapes: listSlideShapes,
  read_slide_text: readSlideText,
  edit_slide_text: editSlideText,
  edit_slide_xml: editSlideXml,
  insert_slide_element: insertSlideElement,
  remove_slide_element: removeSlideElement,
  duplicate_slide: duplicateSlide,
  move_slide: moveSlide,
  modify_presentation_structure: modifyPresentationStructure,
  verify_slides: verifySlides,
  verify_slide_visual: verifySlideVisual,
  export_slide_image: exportSlideImage,
  insert_slide_table: insertSlideTable,
  edit_slide_table: editSlideTable,
  set_slide_background: setSlideBackground,
  set_shape_alt_text: setShapeAltText,
  set_slide_hyperlink: setSlideHyperlink,
  set_shape_zorder: setShapeZOrder,
  group_shapes: groupShapes,
  ungroup_shapes: ungroupShapes,
  align_shapes: alignShapes,
};

function isPptTool(tool: string): tool is PptToolName {
  return Object.prototype.hasOwnProperty.call(HANDLERS, tool);
}

/**
 * A result over the cap. Reads fail so the model asks for less; a write has already happened, so it keeps its summary
 * instead of reporting an error for a change that was made.
 */
export function guardResultSize(data: unknown): unknown {
  if (resultSize(data) <= MAX_RESULT_CHARS) return data;
  const summary = (data as { summary?: unknown } | null)?.summary;
  if (typeof summary === "string") return { summary, note: "The full result was too large to return; read the slide again to see the current state." };
  return capResult(data);
}

/**
 * Execute one PowerPoint tool request inside the open presentation. Arguments were validated by the bridge; they are
 * parsed again here (cheap) so defaults apply and a bad call becomes `invalid_arguments` instead of an exception.
 */
export async function executePptTool(tool: string, args: unknown): Promise<ToolResult> {
  try {
    if (!isPptTool(tool)) throw new UnknownTool(`unknown tool '${tool}'`);
    const parsed = PptToolInputSchemas[tool].safeParse(args ?? {});
    if (!parsed.success) {
      throw new ToolFailure("invalid_arguments", `invalid arguments for ${tool}: ${parsed.error.issues.map((i) => `${i.path.join(".") || "input"}: ${i.message}`).join("; ").slice(0, 400)}`);
    }
    requirePowerPoint();
    const data = guardResultSize(await HANDLERS[tool](parsed.data as never));
    return { ok: true, data };
  } catch (error) {
    if (error instanceof UnknownTool) return { ok: false, error: { code: "unknown_tool", message: error.message } };
    return { ok: false, error: toToolError(error) };
  }
}
