import type { ManageDocPicturesInput, ResolvedInsertImageInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported, type Requirement } from "../../host/requirements.js";
import { fitImage, naturalSize, type Size } from "../image-layout.js";
import { TRACKED_DELETE_NOTE, WORD_API_1_3, assertParagraphIndex, isBlankText, isTracking, preview, requireSupport, wordRun } from "./common.js";
import { deleteParagraph } from "./paragraphs.js";

/** InlinePicture.paragraph and InlinePicture.delete. */
const WORD_API_1_2: Requirement = ["WordApi", "1.2"];
const MAX_LISTED = 200;

/** Without a size a picture is at most this big: the text column of a normal page (6.5 in) and about the height of one. */
const DEFAULT_LIMIT: Size = { width: 468, height: 648 };

/** insert_image in Word: an inline picture in a paragraph of its own. */
export async function insertImage(input: ResolvedInsertImageInput): Promise<unknown> {
  return wordRun(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const after = input.after_paragraph;
    if (after !== undefined) assertParagraphIndex(after, paragraphs.items.length, "after_paragraph");

    // A new paragraph takes the look of its neighbour (a heading, a list item); a picture belongs in a plain one.
    // An empty document already has one blank paragraph: use it instead of leaving it empty above the picture.
    let emptyDocument = false;
    if (paragraphs.items.length === 1) {
      paragraphs.items[0]!.load("text");
      await context.sync();
      emptyDocument = isBlankText(paragraphs.items[0]!.text);
    }
    const holder = emptyDocument && after === undefined
      ? paragraphs.items[0]!
      : after !== undefined ? paragraphs.items[after]!.insertParagraph("", "After") : body.insertParagraph("", "End");
    if (isSupported(WORD_API_1_3)) holder.styleBuiltIn = "Normal";
    const picture = holder.insertInlinePictureFromBase64(input.imageBase64, "Start");
    picture.load("width,height");
    await context.sync();

    const own: Size = picture.width > 0 && picture.height > 0 ? { width: picture.width, height: picture.height } : naturalSize(input) ?? { width: 0, height: 0 };
    const size = fitImage(own, { width: input.width, height: input.height }, DEFAULT_LIMIT);
    picture.lockAspectRatio = false;
    picture.width = size.width;
    picture.height = size.height;
    picture.lockAspectRatio = true;
    if (input.alt_text) picture.altTextDescription = input.alt_text;
    await context.sync();

    const where = after !== undefined ? `after paragraph ${after}` : emptyDocument ? "in the empty document" : "at the end";
    return {
      summary: `Inserted a picture ${where} (${input.sourceLabel})`,
      paragraph_index: after !== undefined ? after + 1 : emptyDocument ? 0 : paragraphs.items.length,
      width: size.width,
      height: size.height,
      source: input.sourceLabel,
      ...(input.name ? { note: "Word pictures have no name; `name` was ignored." } : {}),
    };
  });
}

/** manage_doc_pictures: list, resize / describe, or delete the inline pictures of the body. */
export async function manageDocPictures(input: ManageDocPicturesInput): Promise<unknown> {
  if (input.operation === "delete") requireSupport("Deleting a picture", WORD_API_1_2, "Ask the person to delete it in Word.");
  return wordRun(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("items/width,items/height,items/altTextDescription");
    await context.sync();
    const items = pictures.items;

    if (input.operation === "list") {
      const shown = items.slice(0, MAX_LISTED);
      const withParagraph = isSupported(WORD_API_1_2);
      const paragraphs = withParagraph ? shown.map((picture) => picture.paragraph.load("text")) : [];
      if (withParagraph && shown.length) await context.sync();
      return {
        count: items.length,
        pictures: shown.map((picture, i) => ({
          picture_index: i,
          width: Math.round(picture.width * 100) / 100,
          height: Math.round(picture.height * 100) / 100,
          alt_text: picture.altTextDescription ?? "",
          ...(withParagraph ? { paragraph_text: preview(paragraphs[i]!.text, 80) } : {}),
        })),
        ...(items.length > MAX_LISTED ? { note: `Only the first ${MAX_LISTED} of ${items.length} pictures are listed.` } : {}),
      };
    }

    const index = input.picture_index!;
    const picture = items[index];
    if (!picture) throw new ToolFailure("invalid_arguments", `picture_index ${index} is outside the document, which has ${items.length} picture${items.length === 1 ? "" : "s"}; list them first`);

    if (input.operation === "delete") {
      // A picture alone in its paragraph (as insert_image puts it) takes that paragraph with it, instead of leaving it empty.
      // A table cell keeps its paragraph: deleting it would break the cell.
      const paragraph = picture.paragraph;
      const others = paragraph.inlinePictures;
      paragraph.load(isSupported(WORD_API_1_3) ? "text,tableNestingLevel" : "text");
      others.load("items");
      await context.sync();
      const inTable = isSupported(WORD_API_1_3) && paragraph.tableNestingLevel > 0;
      const alone = isBlankText(paragraph.text) && others.items.length === 1 && !inTable;
      const tracked = await isTracking(context);
      picture.delete();
      if (alone) deleteParagraph(paragraph);
      await context.sync();
      return {
        summary: `Deleted picture ${index}${alone ? " and its empty paragraph" : ""}`,
        deleted: index,
        remaining: items.length - 1,
        ...(alone ? { removed_empty_paragraph: true } : {}),
        ...(tracked ? { tracked: true } : {}),
        note: tracked ? TRACKED_DELETE_NOTE : alone ? "Indexes of later pictures and paragraphs have shifted." : "Indexes of later pictures have shifted.",
      };
    }

    let size: Size = { width: picture.width, height: picture.height };
    if (input.width !== undefined || input.height !== undefined) {
      size = fitImage(size, { width: input.width, height: input.height });
      picture.lockAspectRatio = false;
      picture.width = size.width;
      picture.height = size.height;
      picture.lockAspectRatio = true;
    }
    if (input.alt_text !== undefined) picture.altTextDescription = input.alt_text;
    await context.sync();
    const changed = [input.width !== undefined || input.height !== undefined ? `resized to ${size.width} x ${size.height} pt` : "", input.alt_text !== undefined ? "set its alt text" : ""].filter(Boolean).join(" and ");
    return { summary: `Picture ${index}: ${changed}`, picture_index: index, width: size.width, height: size.height, ...(input.alt_text !== undefined ? { alt_text: input.alt_text } : {}) };
  });
}
