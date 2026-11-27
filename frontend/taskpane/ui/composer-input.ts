/**
 * The message box: one editable element that draws its own text, so a `/command` or `@file` can have a background without a second
 * copy of the text underneath. It offers the part of the textarea interface the pane uses (`value`, selection, `setRangeText`).
 */
export type TextField = Pick<HTMLTextAreaElement, "value" | "selectionStart" | "selectionEnd" | "setSelectionRange" | "setRangeText"
  | "focus" | "dispatchEvent" | "addEventListener" | "removeEventListener" | "closest" | "getClientRects">;
export type ComposerInput = HTMLDivElement & TextField & {
  /** Draw the text again after something besides typing changed what is highlighted. */
  refresh(): void;
  /** Returns the pieces of the text to show (for highlighting). Plain text needs no pieces: the element keeps what the person typed. */
  decorate: (text: string) => Node[];
};

export function createComposerInput(element: HTMLDivElement): ComposerInput {
  const input = element as ComposerInput;
  input.decorate = text => [document.createTextNode(text)];
  let composing = false;
  let lastStart = 0;
  let lastEnd = 0;

  /** The browser adds one newline after a trailing one so the empty last line is visible; it is not part of the text. */
  const read = (): string => {
    const text = input.textContent ?? "";
    return text.endsWith("\n") ? text.slice(0, -1) : text;
  };
  const offsetOf = (node: Node, offset: number): number => {
    const range = document.createRange();
    range.setStart(input, 0);
    range.setEnd(node, offset);
    return Math.min(range.toString().length, read().length);
  };
  const selectionOffsets = (): [number, number] | null => {
    const selection = document.getSelection();
    if (!selection || selection.rangeCount === 0 || !input.contains(selection.anchorNode) || !input.contains(selection.focusNode)) return null;
    const a = offsetOf(selection.anchorNode!, selection.anchorOffset);
    const b = offsetOf(selection.focusNode!, selection.focusOffset);
    return [Math.min(a, b), Math.max(a, b)];
  };
  const place = (start: number, end: number): void => {
    lastStart = start;
    lastEnd = end;
    const locate = (target: number): [Node, number] => {
      const walker = document.createTreeWalker(input, NodeFilter.SHOW_TEXT);
      let left = target;
      let last: Text | null = null;
      for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
        last = node;
        if (left <= node.length) return [node, left];
        left -= node.length;
      }
      return last ? [last, last.length] : [input, 0];
    };
    const [startNode, startOffset] = locate(start);
    const [endNode, endOffset] = locate(end);
    const range = document.createRange();
    range.setStart(startNode, startOffset);
    range.setEnd(endNode, endOffset);
    const selection = document.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  };
  const write = (text: string): void => {
    const nodes = input.decorate(text);
    const highlighted = nodes.some(node => node.nodeType === Node.ELEMENT_NODE);
    // Text with nothing to highlight stays as the browser has it, which keeps undo and the input method working.
    if (!highlighted && !input.querySelector("*")) {
      if (read() !== text) input.textContent = text.endsWith("\n") ? `${text}\n` : text;
      return;
    }
    input.replaceChildren(...nodes, ...(text.endsWith("\n") ? [document.createTextNode("\n")] : []));
  };

  Object.defineProperty(input, "value", {
    configurable: true,
    get: read,
    set(text: string) {
      write(String(text));
      lastStart = lastEnd = read().length;
    },
  });
  Object.defineProperty(input, "selectionStart", { configurable: true, get: () => (selectionOffsets() ?? [lastStart, lastEnd])[0] });
  Object.defineProperty(input, "selectionEnd", { configurable: true, get: () => (selectionOffsets() ?? [lastStart, lastEnd])[1] });
  input.setSelectionRange = (start, end) => place(Math.min(start ?? 0, end ?? 0), Math.max(start ?? 0, end ?? 0));
  input.setRangeText = ((replacement: string, start?: number, end?: number, mode: SelectionMode = "preserve") => {
    const [from, to] = start === undefined ? [input.selectionStart ?? 0, input.selectionEnd ?? 0] : [start, end ?? start];
    const [selStart, selEnd] = [input.selectionStart ?? 0, input.selectionEnd ?? 0];
    const text = read();
    input.value = text.slice(0, from) + replacement + text.slice(to);
    const after = from + replacement.length;
    if (mode === "start") place(from, from);
    else if (mode === "end") place(after, after);
    else if (mode === "select") place(from, after);
    else place(Math.min(selStart, input.value.length), Math.min(selEnd, input.value.length));
  }) as HTMLTextAreaElement["setRangeText"];
  input.refresh = () => {
    if (composing) return;
    const selected = selectionOffsets();
    write(read());
    if (selected) place(selected[0], selected[1]);
  };

  input.addEventListener("compositionstart", () => { composing = true; });
  input.addEventListener("compositionend", () => { composing = false; input.refresh(); });
  input.addEventListener("blur", () => { [lastStart, lastEnd] = selectionOffsets() ?? [lastStart, lastEnd]; });
  return input;
}
