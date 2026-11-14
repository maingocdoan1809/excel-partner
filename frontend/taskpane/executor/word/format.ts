import type { ApplyStyleInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported } from "../../host/requirements.js";
import { assertParagraphIndex, WORD_API_1_3, wordRun } from "./common.js";

const BUILT_IN_STYLES = [
  "Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9",
  "Quote", "IntenseQuote", "ListBullet", "ListNumber", "ListParagraph", "Caption", "NoSpacing", "TocHeading", "Header", "Footer", "Strong", "Emphasis",
] as const;

/** `Heading 1`, `heading1` and `Heading1` all name the built-in style; anything else is a custom style name. */
export function builtInStyleId(name: string): string | undefined {
  const key = name.replace(/[\s_-]+/g, "").toLowerCase();
  return BUILT_IN_STYLES.find((id) => id.toLowerCase() === key);
}

type Styled = Word.Paragraph | Word.Range;

function setStyle(target: Styled, style: string): void {
  const builtIn = builtInStyleId(style);
  // The built-in id is language independent; the name is only used when the host cannot take the id.
  if (builtIn && isSupported(WORD_API_1_3)) (target as { styleBuiltIn: string }).styleBuiltIn = builtIn;
  else target.style = style;
}

function setFont(target: Styled, font: NonNullable<ApplyStyleInput["font"]>): void {
  const f = target.font;
  if (font.name !== undefined) f.name = font.name;
  if (font.size !== undefined) f.size = font.size;
  if (font.bold !== undefined) f.bold = font.bold;
  if (font.italic !== undefined) f.italic = font.italic;
  if (font.underline !== undefined) f.underline = font.underline;
  if (font.color !== undefined) f.color = font.color;
  if (font.highlight !== undefined) f.highlightColor = font.highlight as string;
}

function setParagraphProps(paragraph: Word.Paragraph, input: ApplyStyleInput): void {
  if (input.alignment !== undefined) paragraph.alignment = input.alignment;
  if (input.space_after !== undefined) paragraph.spaceAfter = input.space_after;
  if (input.space_before !== undefined) paragraph.spaceBefore = input.space_before;
  if (input.line_spacing !== undefined) paragraph.lineSpacing = input.line_spacing;
}

export async function applyStyle(input: ApplyStyleInput): Promise<unknown> {
  const target = input.target ?? "selection";
  const needsParagraphs = input.alignment !== undefined || input.space_after !== undefined || input.space_before !== undefined || input.line_spacing !== undefined;
  return wordRun(async (context) => {
    let affected = 0;
    if (target === "selection") {
      const selection = context.document.getSelection();
      selection.load("text");
      const paragraphs = selection.paragraphs;
      if (needsParagraphs) paragraphs.load("items");
      await context.sync();
      if (selection.text === "" && input.font) {
        throw new ToolFailure("invalid_arguments", "The selection is only a cursor, so a font change would affect nothing. Select the text first, or use target=paragraphs with paragraph_index.");
      }
      if (input.style !== undefined) setStyle(selection, input.style);
      if (input.font) setFont(selection, input.font);
      if (needsParagraphs) for (const p of paragraphs.items) setParagraphProps(p, input);
      affected = needsParagraphs ? paragraphs.items.length : 1;
      await context.sync();
    } else {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      const start = input.paragraph_index!;
      assertParagraphIndex(start, paragraphs.items.length);
      const chosen = paragraphs.items.slice(start, start + (input.count ?? 1));
      for (const p of chosen) {
        if (input.style !== undefined) setStyle(p, input.style);
        if (input.font) setFont(p, input.font);
        setParagraphProps(p, input);
      }
      affected = chosen.length;
      await context.sync();
    }
    const applied = [
      input.style !== undefined ? `style ${input.style}` : undefined,
      input.font ? `font (${Object.keys(input.font).join(", ")})` : undefined,
      input.alignment !== undefined ? `alignment ${input.alignment}` : undefined,
      input.space_before !== undefined || input.space_after !== undefined || input.line_spacing !== undefined ? "spacing" : undefined,
    ].filter((x): x is string => x !== undefined);
    return {
      summary: `Formatted ${target === "selection" ? "the selection" : `${affected} paragraph${affected === 1 ? "" : "s"} from ${input.paragraph_index}`}: ${applied.join(", ")}`,
      target,
      ...(target === "paragraphs" ? { paragraph_index: input.paragraph_index, paragraphs: affected } : { paragraphs: affected }),
      applied,
    };
  });
}
