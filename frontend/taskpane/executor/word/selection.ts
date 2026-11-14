import type { WordSelectionSnapshot } from "../../../../protocol/src/index.js";
import { isSupported, WORD_COMMENTS } from "../../host/requirements.js";
import { clip, normalizeNewlines, WORD_TRACK_MODE, wordRun } from "./common.js";

const SELECTION_TEXT_LIMIT = 4000;
/** Above this many characters the paragraph count comes from the text instead of loading every paragraph. */
const PARAGRAPH_LOAD_LIMIT = 20_000;

interface SelectionCore {
  text: string;
  style: string | undefined;
  paragraphs: number;
  highlight: string | null;
}

async function readCore(context: Word.RequestContext, withHighlight: boolean): Promise<{ core: SelectionCore; selection: Word.Range }> {
  const selection = context.document.getSelection();
  const font = selection.font;
  selection.load("text,style");
  if (withHighlight) font.load("highlightColor");
  await context.sync();
  const text = selection.text ?? "";
  let paragraphs: number;
  if (text.length <= PARAGRAPH_LOAD_LIMIT) {
    const list = selection.paragraphs;
    list.load("items");
    await context.sync();
    paragraphs = Math.max(1, list.items.length);
  } else {
    paragraphs = (text.match(/\r/g)?.length ?? 0) + (text.endsWith("\r") ? 0 : 1);
  }
  const color = withHighlight ? font.highlightColor : null;
  return { selection, core: { text, style: selection.style || undefined, paragraphs, highlight: color === undefined || color === "" ? null : color } };
}

/** What the person has selected in Word. Selected text is data from the document, not instructions. */
export async function getWordSelection(): Promise<WordSelectionSnapshot> {
  return wordRun(async (context) => {
    const { core, selection } = await readCore(context, true);
    const text = clip(normalizeNewlines(core.text), SELECTION_TEXT_LIMIT);
    const snapshot: WordSelectionSnapshot = {
      host: "word",
      text: text.text,
      cursor: core.text.length === 0,
      ...(core.style ? { style: core.style } : {}),
      paragraphs: core.paragraphs,
      highlight: core.highlight,
      truncated: text.truncated,
    };

    // Comments and the Track Changes flag are extras: a host that cannot give them must not lose the selection itself.
    if (isSupported(WORD_COMMENTS)) {
      try {
        const comments = selection.getComments();
        comments.load("items/id,items/content,items/resolved");
        await context.sync();
        const withReplies = comments.items.slice(0, 20).map((comment) => {
          const replies = comment.replies;
          replies.load("items");
          return { comment, replies };
        });
        if (withReplies.length > 0) await context.sync();
        snapshot.comments = withReplies.map(({ comment: c, replies }) => ({ id: c.id, text: clip(c.content, 300).text, replies: replies.items.length, resolved: c.resolved }));
      } catch {
        // leave comments out
      }
    }
    if (isSupported(WORD_TRACK_MODE)) {
      try {
        context.document.load("changeTrackingMode");
        await context.sync();
        snapshot.trackChanges = context.document.changeTrackingMode !== "Off";
      } catch {
        // leave the flag out
      }
    }
    return snapshot;
  });
}

/** Short text for the selection chip and a longer one for its tooltip. */
export async function getWordSelectionLabel(): Promise<{ label: string; title: string }> {
  return wordRun(async (context) => {
    const { core } = await readCore(context, false);
    const style = core.style;
    if (core.text.length === 0) {
      return { label: style ? `Cursor in ${style}` : "Cursor", title: style ? `Cursor in a paragraph styled “${style}”` : "Cursor (no text selected)" };
    }
    const flat = normalizeNewlines(core.text).replace(/\s+/g, " ").trim();
    const shortText = flat.length > 40 ? `${flat.slice(0, 40)}…` : flat;
    const count = core.paragraphs > 1 ? ` · ${core.paragraphs} paragraphs` : "";
    return {
      label: `“${shortText}”${count}`,
      title: `${clip(flat, 300).text}\n${core.text.length} characters · ${core.paragraphs} paragraph${core.paragraphs === 1 ? "" : "s"}${style ? ` · ${style}` : ""}`,
    };
  });
}

/** File name from the document URL. Word has no reliable name API, so a saved-as or renamed file is not noticed. */
export async function readWordDocumentName(): Promise<string | null> {
  try {
    const url = Office.context.document.url;
    if (!url) return null;
    const last = url.split(/[\\/]/).pop() ?? "";
    const base = last.split(/[?#]/)[0] ?? "";
    let name = base;
    try {
      name = decodeURIComponent(base);
    } catch {
      // keep the raw name
    }
    return name.trim() === "" ? null : name;
  } catch {
    return null;
  }
}
