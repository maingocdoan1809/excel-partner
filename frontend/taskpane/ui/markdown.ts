/// <reference lib="dom" />
import DOMPurify from "dompurify";
import { Marked } from "marked";
import { parseCellHref } from "../../../protocol/src/index.js";
import { cellBadge, linkifyCells } from "./cell-links.js";

const markdown = new Marked({ gfm: true, breaks: true });

export interface MarkdownOptions {
  /**
   * Cell and range links that select them in Excel. A markdown link to `cell:?file=…&sheet=…&range=…` always becomes a
   * badge that goes exactly there. References written in plain words are found too unless `plain` is false (a reply about
   * another file: its bare cells could belong to that file, so none is guessed).
   */
  cellLinks?: boolean | { plain?: boolean };
}

/** Render GFM (including tables) and sanitize provider-supplied HTML. */
export function renderMarkdown(source: string, options: MarkdownOptions = {}): string {
  const html = markdown.parse(source, { async: false });
  const safe = DOMPurify.sanitize(html, {
    ALLOWED_TAGS: ["p", "br", "strong", "em", "del", "code", "pre", "blockquote", "ul", "ol", "li",
      "h1", "h2", "h3", "h4", "h5", "h6", "hr", "a", "table", "thead", "tbody", "tfoot", "tr", "th", "td"],
    ALLOWED_ATTR: ["href", "title", "colspan", "rowspan", "align"],
    // Web links, and our own cell links; anything else loses its address and is shown as text.
    ALLOWED_URI_REGEXP: /^(?:https?|cell):/i,
  });
  const container = document.createElement("div");
  container.innerHTML = safe;
  for (const link of Array.from(container.querySelectorAll("a"))) {
    const href = link.getAttribute("href") ?? "";
    if (/^cell:/i.test(href)) {
      const target = options.cellLinks ? parseCellHref(href) : null;
      // One that names no real range is only text: the words the assistant wrote, with nothing to click.
      if (target) link.replaceWith(cellBadge(link.textContent?.trim() || target.address, target));
      else link.replaceWith(...Array.from(link.childNodes));
      continue;
    }
    if (!/^https?:\/\//i.test(href)) {
      link.replaceWith(...Array.from(link.childNodes));
      continue;
    }
    link.target = "_blank";
    link.rel = "noopener noreferrer";
  }
  for (const table of Array.from(container.querySelectorAll("table"))) {
    const wrapper = document.createElement("div");
    wrapper.className = "markdown-table";
    table.replaceWith(wrapper);
    wrapper.append(table);
  }
  if (options.cellLinks && !(typeof options.cellLinks === "object" && options.cellLinks.plain === false)) linkifyCells(container);
  return container.innerHTML;
}
