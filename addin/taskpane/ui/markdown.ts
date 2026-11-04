/// <reference lib="dom" />
import DOMPurify from "dompurify";
import { Marked } from "marked";
import { linkifyCells } from "./cell-links.js";

const markdown = new Marked({ gfm: true, breaks: true });

/** Render GFM (including tables) and sanitize provider-supplied HTML. With `cellLinks`, cell and range mentions become links that select them in Excel. */
export function renderMarkdown(source: string, options: { cellLinks?: boolean } = {}): string {
  const html = markdown.parse(source, { async: false });
  const safe = DOMPurify.sanitize(html, {
    ALLOWED_TAGS: ["p", "br", "strong", "em", "del", "code", "pre", "blockquote", "ul", "ol", "li",
      "h1", "h2", "h3", "h4", "h5", "h6", "hr", "a", "table", "thead", "tbody", "tfoot", "tr", "th", "td"],
    ALLOWED_ATTR: ["href", "title", "colspan", "rowspan", "align"],
  });
  const container = document.createElement("div");
  container.innerHTML = safe;
  for (const link of Array.from(container.querySelectorAll("a"))) {
    if (!/^https?:\/\//i.test(link.getAttribute("href") ?? "")) {
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
  if (options.cellLinks) linkifyCells(container);
  return container.innerHTML;
}
