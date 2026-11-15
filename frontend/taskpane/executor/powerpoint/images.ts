import type { ResolvedInsertImageInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { centreIn, fitImage, naturalSize } from "../image-layout.js";
import { getSlideSize } from "./ooxml.js";
import { selectedSlideId } from "./selection.js";
import { PPT_MODERN, isSupported, makeRef, requireSlide, requireSupport, round, type Requirement } from "./util.js";

/** The picture API of the Common Office API: the only way PowerPoint takes a picture in from base64. */
const IMAGE_COERCION: Requirement = ["ImageCoercion", "1.2"];

function officeCall(what: string, start: (done: (result: Office.AsyncResult<unknown>) => void) => void): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      start((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new ToolFailure("office_error", `${what}: ${result.error?.message ?? "PowerPoint refused"}`));
      });
    } catch (error) {
      reject(new ToolFailure("office_error", `${what}: ${error instanceof Error ? error.message : String(error)}`));
    }
  });
}

/** Show a slide, because pictures go onto the slide the person is looking at. */
function showSlide(slideId: string): Promise<void> {
  const numeric = Number.parseInt(slideId, 10);
  if (!Number.isFinite(numeric)) throw new ToolFailure("invalid_arguments", `slide id '${slideId}' is not one PowerPoint can show; read the current ids with get_presentation_metadata`);
  return officeCall("could not open the slide", (done) => Office.context.document.goToByIdAsync(numeric, Office.GoToType.Slide, done));
}

/** insert_image in PowerPoint: a picture on a slide. */
export async function insertImage(input: ResolvedInsertImageInput): Promise<unknown> {
  requireSupport("insert_image", IMAGE_COERCION, "Ask the person to insert the picture by hand (Insert > Pictures), or to update Office.");
  return PowerPoint.run(async (context) => {
    const slideId = input.slide_id ?? await selectedSlideId(context);
    const slide = await requireSlide(context, slideId);
    slide.shapes.load("items/id");
    const slideSize = await getSlideSize(context, slideId);
    await context.sync();
    const before = new Set(slide.shapes.items.map((shape) => shape.id));

    // The size is set while inserting, so it has to be worked out from the pixel size the bridge read.
    const own = naturalSize(input);
    const requested = input.width !== undefined || input.height !== undefined;
    const fitted = own ? fitImage(own, { width: input.width, height: input.height }, slideSize) : null;
    if (!fitted && requested) throw new ToolFailure("invalid_arguments", "this picture does not say how big it is, so width and height cannot be applied; leave them out");
    const place = fitted
      ? (input.left !== undefined || input.top !== undefined || !slideSize
        ? { ...fitted, left: input.left ?? 0, top: input.top ?? 0 }
        : centreIn({ left: 0, top: 0, width: slideSize.width, height: slideSize.height }, fitted))
      : null;

    await showSlide(slideId);
    await officeCall("PowerPoint did not take the picture", (done) => Office.context.document.setSelectedDataAsync(input.imageBase64, {
      coercionType: Office.CoercionType.Image,
      ...(place ? { imageLeft: place.left, imageTop: place.top, imageWidth: place.width, imageHeight: place.height } : {}),
    }, done));

    slide.shapes.load("items/id");
    await context.sync();
    const added = slide.shapes.items.filter((shape) => !before.has(shape.id));
    const picture = added[added.length - 1];
    if (!picture) throw new ToolFailure("office_error", "PowerPoint reported success but no picture appeared on the slide; check it with list_slide_shapes before trying again.");
    if (input.name) picture.name = input.name;
    let note: string | undefined;
    if (input.alt_text) {
      if (isSupported(PPT_MODERN)) picture.altTextDescription = input.alt_text;
      else note = "The picture was inserted, but alt text needs PowerPointApi 1.10; set it by hand or with set_shape_alt_text.";
    }
    picture.load("id,name,left,top,width,height");
    await context.sync();
    return {
      summary: `Added a picture to slide ${slideId} ("${picture.name}", ${input.sourceLabel})`,
      slide_id: slideId,
      ref: makeRef(slideId, picture.id),
      shape_id: picture.id,
      name: picture.name,
      left: round(picture.left),
      top: round(picture.top),
      width: round(picture.width),
      height: round(picture.height),
      source: input.sourceLabel,
      ...(added.length > 1 ? { warning: `${added.length} shapes appeared on the slide; ref points at the last one, check list_slide_shapes.` } : {}),
      ...(note ? { note } : {}),
    };
  });
}
