import { isSupported } from "../../host/requirements.js";
import { WORD_API_1_3 } from "./common.js";

export interface ParagraphInfo {
  /** 0-based position in `body.paragraphs`. */
  index: number;
  text: string;
  style: string;
  /** 1..9 for headings (a Title counts as 1), null for body text. */
  headingLevel: number | null;
  /** Inside a table cell (only known with WordApi 1.3). */
  inTable: boolean;
  /** Part of a bulleted / numbered list (only known with WordApi 1.3). */
  listItem: boolean;
}

export interface LoadedParagraphs {
  collection: Word.ParagraphCollection;
  info: ParagraphInfo[];
}

const HEADING_STYLE = /^(?:heading|tiêu đề|titre|überschrift|encabezado|título|intestazione|kop|rubrik|заголовок|見出し|제목)\s*(\d)$/iu;

export function headingLevelOf(style: string, builtIn: string | undefined): number | null {
  const built = /^Heading([1-9])$/.exec(builtIn ?? "");
  if (built) return Number(built[1]);
  if (builtIn === "Title") return 1;
  const named = HEADING_STYLE.exec(style.trim());
  if (named) return Number(named[1]);
  if (/^title$/i.test(style.trim())) return 1;
  return null;
}

/** Load every body paragraph with the properties the read tools need. Only asks for what this Word version has. */
export async function loadParagraphs(context: Word.RequestContext, options: { structure?: boolean } = {}): Promise<LoadedParagraphs> {
  const collection = context.document.body.paragraphs;
  const modern = isSupported(WORD_API_1_3);
  const names = ["items/text", "items/style"];
  if (modern) names.push("items/styleBuiltIn");
  if (modern && options.structure !== false) names.push("items/tableNestingLevel", "items/isListItem");
  collection.load(names.join(","));
  await context.sync();
  const info = collection.items.map((p, index): ParagraphInfo => {
    const style = p.style ?? "";
    return {
      index,
      text: p.text,
      style,
      headingLevel: headingLevelOf(style, modern ? (p.styleBuiltIn as string | undefined) : undefined),
      inTable: modern && options.structure !== false ? p.tableNestingLevel > 0 : false,
      listItem: modern && options.structure !== false ? p.isListItem : false,
    };
  });
  return { collection, info };
}

/** The paragraph's own range, for deleting or anchoring (`Paragraph` has no delete in the contract the plan relies on). */
export function paragraphRange(paragraph: Word.Paragraph, location: "Whole" | "Start" | "End" | "After" = "Whole"): Word.Range {
  return paragraph.getRange(location);
}

/** Delete a paragraph through its range. */
export function deleteParagraph(paragraph: Word.Paragraph): void {
  if (typeof paragraph.getRange === "function") paragraph.getRange("Whole").delete();
  else paragraph.delete();
}
