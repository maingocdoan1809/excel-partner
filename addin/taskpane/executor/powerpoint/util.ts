import { ToolFailure } from "../common.js";
import { isSupported, requireSupport, PPT_EXPORT, PPT_MODERN, PPT_SELECTION, type Requirement } from "../../host/requirements.js";

export { isSupported, requireSupport, PPT_EXPORT, PPT_MODERN, PPT_SELECTION };
export type { Requirement };

const api = (version: string): Requirement => ["PowerPointApi", version];

/** API sets the PowerPoint tools need beyond the 1.2 baseline (numbers checked against @types/office-js). */
export const PPT_LAYOUTS = api("1.3"); // slideMasters, layouts, slides.add, slide.layout, shape.delete
export const PPT_GEOMETRY = api("1.4"); // shape left/top/width/height, textFrame, fill, lineFormat, addTextBox/addGeometricShape/addLine
export const PPT_ZORDER = api("1.8"); // setZOrder, addGroup, group.ungroup, placeholderFormat
export const PPT_TABLES = api("1.8"); // addTable, getTable, Table.values
export const PPT_TABLE_EDIT = api("1.9"); // rows/columns add + delete, mergeCells
export const PPT_MOVE = api("1.8"); // Slide.moveTo / index

/** Largest JSON result (characters) sent back to the model. */
export const MAX_RESULT_CHARS = 30_000;
/** Largest image (base64 characters) returned to the model. */
export const MAX_IMAGE_BASE64 = 1_500_000;

export function slideNotFound(slideId: string, known?: string[]): ToolFailure {
  return new ToolFailure(
    "invalid_arguments",
    `slide '${slideId}' does not exist. Slide ids change when a slide is rebuilt by edit_slide_xml, duplicate_slide or add_slide; read the current ids with get_presentation_metadata.`,
    known ? { slide_ids: known.slice(0, 200) } : undefined,
  );
}

export function parseShapeRef(ref: string): { slideId: string; shapeId: string } {
  const match = /^([^:\s]+):([^:\s]+)$/.exec(ref ?? "");
  if (!match) throw new ToolFailure("invalid_arguments", `shape ref '${String(ref)}' must look like "<slideId>:<shapeId>", e.g. "256:4" (see list_slide_shapes)`);
  return { slideId: match[1]!, shapeId: match[2]! };
}

export const makeRef = (slideId: string, shapeId: string): string => `${slideId}:${shapeId}`;

/** Ids of all slides in presentation order. */
export async function loadSlideIds(context: PowerPoint.RequestContext): Promise<string[]> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  return slides.items.map((slide) => slide.id);
}

/** The slide with this id, or a ToolFailure that lists the real ids. */
export async function requireSlide(context: PowerPoint.RequestContext, slideId: string): Promise<PowerPoint.Slide> {
  const slide = context.presentation.slides.getItemOrNullObject(slideId);
  slide.load("isNullObject");
  await context.sync();
  if (slide.isNullObject) throw slideNotFound(slideId, await loadSlideIds(context));
  return slide;
}

/** A shape of a slide, or a ToolFailure. `shapeId` is the second half of a ref. */
export async function requireShape(context: PowerPoint.RequestContext, slideId: string, shapeId: string): Promise<{ slide: PowerPoint.Slide; shape: PowerPoint.Shape }> {
  const slide = await requireSlide(context, slideId);
  const shape = slide.shapes.getItemOrNullObject(shapeId);
  shape.load("isNullObject");
  await context.sync();
  if (shape.isNullObject) {
    throw new ToolFailure("invalid_arguments", `shape '${makeRef(slideId, shapeId)}' does not exist; list the shapes of the slide with list_slide_shapes`);
  }
  return { slide, shape };
}

/** Normalise PowerPoint line separators (\r paragraph, \v soft break) to \n. */
export function normalizeText(text: string | null | undefined): string {
  return (text ?? "").replace(/\r\n|\r|\u000b/g, "\n");
}

export function clip(text: string, max: number): { text: string; truncated: boolean } {
  return text.length > max ? { text: text.slice(0, max), truncated: true } : { text, truncated: false };
}

export function round(value: number, digits = 1): number {
  const k = 10 ** digits;
  return Math.round(value * k) / k;
}

/** JSON size of a result, not counting a rendered image (which has its own cap). */
export function resultSize(data: unknown): number {
  if (data && typeof data === "object" && "image" in (data as Record<string, unknown>)) {
    const { image: _image, ...rest } = data as Record<string, unknown>;
    return JSON.stringify(rest).length;
  }
  return JSON.stringify(data ?? null).length;
}

/** Enforce the 30,000 character cap on a tool result. */
export function capResult<T>(data: T): T {
  const size = resultSize(data);
  if (size > MAX_RESULT_CHARS) {
    throw new ToolFailure("size_limit_exceeded", `the result is ${size} characters (limit ${MAX_RESULT_CHARS}); ask for fewer slides or shapes`);
  }
  return data;
}

/** Highest PowerPointApi version this host reports, for the model's benefit. */
export function highestPptVersion(): string | null {
  for (const version of ["1.10", "1.9", "1.8", "1.7", "1.6", "1.5", "1.4", "1.3", "1.2", "1.1"]) {
    if (isSupported(api(version))) return version;
  }
  return null;
}

export function requirePowerPoint(): void {
  if (typeof PowerPoint === "undefined" || typeof PowerPoint.run !== "function") {
    throw new ToolFailure("unsupported", "PowerPoint.run is not available in this task pane; open the add-in inside PowerPoint.");
  }
}
