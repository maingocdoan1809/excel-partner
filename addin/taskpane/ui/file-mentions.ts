export interface ActiveFileMention { start: number; query: string }

// Characters that make an @ part of a word (an e-mail address such as a.b@c.com), not a mention.
const WORD_CHARACTER = /[\p{L}\p{N}_.+%-]/u;

/** An @ opens a mention at the start of the text or after anything that is not part of a word. */
export function isMentionBoundary(previous: string | undefined): boolean {
  return previous === undefined || !WORD_CHARACTER.test(previous);
}

/**
 * Find the mention being typed. It may sit anywhere in the prompt, after punctuation such as "(",
 * and a second @ (typed on its own or right after the first) keeps the menu open.
 */
export function activeFileMention(beforeCursor: string, selectedNames: readonly string[]): ActiveFileMention | null {
  const ranges = fileMentionRanges(beforeCursor, selectedNames);
  for (let at = beforeCursor.length - 1; at >= 0; at--) {
    if (beforeCursor[at] !== "@") continue;
    // "@@" is one trigger: the whole run is replaced when a file is chosen.
    let start = at;
    while (start > 0 && beforeCursor[start - 1] === "@") start--;
    // Not a mention here (e-mail address): keep looking at the @ before it.
    if (!isMentionBoundary(beforeCursor[start - 1])) continue;
    if (ranges.some((range) => start >= range.start && start < range.end)) return null;
    const query = beforeCursor.slice(at + 1);
    if (query.includes("\n")) return null;
    if (selectedNames.some((name) => query === name || (query.startsWith(name) && /^\s/.test(query.slice(name.length))))) return null;
    return { start, query };
  }
  return null;
}

export interface FileMentionRange { start: number; end: number; name: string }

/** Exact selected mentions to decorate in the composer and sent user bubble. */
export function fileMentionRanges(text: string, selectedNames: readonly string[]): FileMentionRange[] {
  const names = [...new Set(selectedNames)].sort((a, b) => b.length - a.length);
  const ranges: FileMentionRange[] = [];
  for (let i = 0; i < text.length; i++) {
    if (text[i] !== "@" || !isMentionBoundary(text[i - 1])) continue;
    const name = names.find((item) => text.startsWith(item, i + 1) &&
      (i + 1 + item.length === text.length || /[\s,;.!?)}\]]/.test(text[i + 1 + item.length]!)));
    if (!name) continue;
    ranges.push({ start: i, end: i + 1 + name.length, name });
    i += name.length;
  }
  return ranges;
}

/** Expand a deletion touching any badge to remove the whole token. */
export function atomicBadgeDeletion(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  inputType: string,
  badges: ReadonlyArray<{ start: number; end: number }>,
): { start: number; end: number } | null {
  if (!badges.length) return null;
  let start = selectionStart;
  let end = selectionEnd;
  if (start === end) {
    const backwards = inputType.toLowerCase().includes("backward");
    const position = backwards ? start - 1 : start;
    const touched = badges.find((mention) => mention.start <= position && position < mention.end);
    if (touched) return { start: touched.start, end: touched.start === 0 && text[touched.end] === " " ? touched.end + 1 : touched.end };
    if (backwards && text[position] === " ") {
      const preceding = badges.find((mention) => mention.end === position);
      if (preceding) return { start: preceding.start, end: position + 1 };
    }
    return null;
  }
  const touched = badges.filter((mention) => mention.start < end && mention.end > start);
  if (!touched.length) return null;
  for (const mention of touched) {
    start = Math.min(start, mention.start);
    end = Math.max(end, mention.end);
  }
  if (text[end] === " " && (start === 0 || text[start - 1] === " ")) end++;
  return { start, end };
}

/** Compatibility wrapper for file-only deletion checks. */
export function atomicFileMentionDeletion(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  inputType: string,
  selectedNames: readonly string[],
): { start: number; end: number } | null {
  return atomicBadgeDeletion(text, selectionStart, selectionEnd, inputType, fileMentionRanges(text, selectedNames));
}
