/**
 * Tiny, safe markdown for chat bubbles: links, **bold**, *italic*, `code`, fenced
 * code blocks, `-`/`1.` lists and line breaks. Everything is HTML-escaped
 * first, so the result is safe to assign to `innerHTML`.
 */
export function renderMarkdown(source: string): string {
  const escaped = source.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const out: string[] = [];
  let list: "ul" | "ol" | null = null;
  let fence: string[] | null = null;

  const closeList = () => {
    if (list) out.push(`</${list}>`);
    list = null;
  };

  for (const line of escaped.split("\n")) {
    if (line.trimStart().startsWith("```")) {
      if (fence) {
        out.push(`<pre><code>${fence.join("\n")}</code></pre>`);
        fence = null;
      } else {
        closeList();
        fence = [];
      }
      continue;
    }
    if (fence) {
      fence.push(line);
      continue;
    }

    const bullet = /^\s*[-*]\s+(.*)$/.exec(line);
    const numbered = /^\s*\d+[.)]\s+(.*)$/.exec(line);
    const item = bullet ?? numbered;
    if (item) {
      const kind = bullet ? "ul" : "ol";
      if (list !== kind) {
        closeList();
        out.push(`<${kind}>`);
        list = kind;
      }
      out.push(`<li>${inline(item[1]!)}</li>`);
      continue;
    }

    closeList();
    out.push(line.trim() === "" ? "<br>" : `<p>${inline(line)}</p>`);
  }
  if (fence) out.push(`<pre><code>${fence.join("\n")}</code></pre>`); // unterminated while streaming
  closeList();
  return out.join("");
}

function inline(text: string): string {
  const tokens = /`([^`]+)`|\[([^\]\n]+)\]\((https?:\/\/(?:[^\s()]|\([^\s()]*\))+)\)/g;
  let result = "";
  let offset = 0;
  for (const match of text.matchAll(tokens)) {
    result += emphasis(text.slice(offset, match.index));
    if (match[1] !== undefined) result += `<code>${match[1]}</code>`;
    else {
      const href = match[3]!.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
      result += `<a href="${href}" target="_blank" rel="noopener noreferrer">${emphasis(match[2]!)}</a>`;
    }
    offset = match.index! + match[0].length;
  }
  return result + emphasis(text.slice(offset));
}

function emphasis(text: string): string {
  return text
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>");
}
