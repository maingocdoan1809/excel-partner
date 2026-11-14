import type { EditSlideXmlInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { NS, makeElementHelper, makeQueryHelper, roundTripSlide } from "./ooxml.js";
import { PPT_EXPORT, requireSupport } from "./util.js";

/** Largest JSON the model's code may return; leaves room for the rest of the result under the 30,000 character cap. */
const MAX_RETURNED_CHARS = 25_000;

const AsyncFunction = Object.getPrototypeOf(async function () {
  /* constructor access only */
}).constructor as new (...args: string[]) => (...values: unknown[]) => Promise<unknown>;

function capReturned(value: unknown): unknown {
  if (value === undefined) return null;
  let json: string;
  try {
    json = JSON.stringify(value);
  } catch {
    return "(the returned value is not JSON-serialisable; return plain strings, numbers or objects)";
  }
  if (json === undefined || json.length <= MAX_RETURNED_CHARS) return value;
  return { truncated: true, characters: json.length, preview: json.slice(0, 2000) };
}

/**
 * Change a slide's XML with code the model wrote. The slide is exported (PowerPointApi 1.8), `xml` is the parsed
 * ppt/slides/slide1.xml, the code edits it in place, and the result is inserted back where the slide was. The original is
 * deleted only after the new slide is confirmed in the deck. The new slide has a new id.
 */
export async function editSlideXml(input: EditSlideXmlInput): Promise<unknown> {
  requireSupport("edit_slide_xml", PPT_EXPORT, "Editing slide XML goes through the slide's exported .pptx, which needs PowerPointApi 1.8. Use edit_slide_text, insert_slide_element or execute_office_js instead.");
  let fn: (...values: unknown[]) => Promise<unknown>;
  try {
    fn = new AsyncFunction("xml", "files", "ns", "helpers", input.code);
  } catch (error) {
    throw new ToolFailure("invalid_arguments", `the code does not compile: ${error instanceof Error ? error.message : String(error)}`);
  }
  return PowerPoint.run(async (context) => {
    const result = await roundTripSlide(context, {
      slideId: input.slide_id,
      mode: "replace",
      mutate: async (doc, parts) => {
        const helpers = { el: makeElementHelper(doc), q: makeQueryHelper(doc) };
        try {
          return capReturned(await fn(doc, Object.keys(parts.files), { ...NS }, helpers));
        } catch (error) {
          throw new ToolFailure("invalid_arguments", `your code threw: ${error instanceof Error ? error.message : String(error)}. The slide is unchanged.`);
        }
      },
    });
    return {
      summary: `Edited the XML of slide ${result.position}${input.description ? `: ${input.description}` : ""} (new slide id ${result.newSlideId})`,
      slide_id: result.newSlideId,
      replaced_slide_id: result.replacedSlideId,
      position: result.position,
      slide_count: result.slideCount,
      shape_count: result.shapeCount,
      result: result.returned ?? null,
      note: "The slide was rebuilt through its .pptx: it has a new slide id and the old id is gone. Shape ids inside it are kept. Check it with verify_slides.",
      ...(result.warnings.length ? { warnings: result.warnings } : {}),
    };
  });
}
