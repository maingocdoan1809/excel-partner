/// <reference lib="dom" />
import { renderMarkdown } from "./markdown.js";

/** A file the assistant wants the user to look at. It lives in the chat history, never on disk. */
export interface StoredArtifact {
  name: string;
  content: string;
}

export type ArtifactKind = "html" | "svg" | "markdown" | "csv" | "json" | "text";

const KIND_BY_EXTENSION: Record<string, ArtifactKind> = {
  html: "html", htm: "html", svg: "svg", md: "markdown", markdown: "markdown", csv: "csv", json: "json", txt: "text",
};
const KIND_LABEL: Record<ArtifactKind, string> = {
  html: "HTML page", svg: "SVG image", markdown: "Markdown", csv: "CSV table", json: "JSON", text: "Text",
};

export function artifactKind(name: string): ArtifactKind {
  return KIND_BY_EXTENSION[name.slice(name.lastIndexOf(".") + 1).toLowerCase()] ?? "text";
}

/**
 * The preview runs in a sandboxed iframe without `allow-same-origin`, so it cannot read the task pane, the
 * workbook or its storage. This policy also closes the network: only inline code, data: images and two script CDNs.
 */
const HTML_POLICY = "default-src 'none'; script-src 'unsafe-inline' https://cdnjs.cloudflare.com https://cdn.jsdelivr.net; "
  + "style-src 'unsafe-inline'; img-src data: blob:; font-src data:; form-action 'none'";
const STATIC_POLICY = "default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:";

const PAGE_STYLE = `
:root { color-scheme: light dark; }
body { margin: 0; padding: 16px; font: 14px/1.55 system-ui, "Segoe UI", sans-serif; background: Canvas; color: CanvasText; }
pre { margin: 0; white-space: pre-wrap; word-break: break-word; font: 12px/1.5 Consolas, monospace; }
table { border-collapse: collapse; }
th, td { border: 1px solid color-mix(in srgb, CanvasText 25%, transparent); padding: 4px 8px; text-align: left; vertical-align: top; }
th { background: color-mix(in srgb, CanvasText 8%, transparent); }
code { font-family: Consolas, monospace; }
img, svg { max-width: 100%; height: auto; }`;

function escapeHtml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** RFC 4180 style: quoted fields may hold commas, quotes ("") and line breaks. */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i]!;
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (char === '"') quoted = false;
      else field += char;
    } else if (char === '"') quoted = true;
    else if (char === "," || (char === ";" && !text.includes(","))) { row.push(field); field = ""; }
    else if (char === "\n" || char === "\r") {
      if (char === "\r" && text[i + 1] === "\n") i++;
      row.push(field); field = "";
      rows.push(row); row = [];
    } else field += char;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function csvTable(text: string): string {
  const rows = parseCsv(text);
  if (!rows.length) return "<p>Empty file.</p>";
  const cells = (row: string[], tag: string) => `<tr>${row.map(cell => `<${tag}>${escapeHtml(cell)}</${tag}>`).join("")}</tr>`;
  return `<table><thead>${cells(rows[0]!, "th")}</thead><tbody>${rows.slice(1).map(row => cells(row, "td")).join("")}</tbody></table>`;
}

function prettyJson(text: string): string {
  try { return JSON.stringify(JSON.parse(text), null, 2); } catch { return text; }
}

function shell(policy: string, body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${policy}">`
    + `<meta name="viewport" content="width=device-width, initial-scale=1"><style>${PAGE_STYLE}</style></head><body>${body}</body></html>`;
}

/** Full document for the preview iframe's `srcdoc`. */
export function buildPreviewDocument(artifact: StoredArtifact): string {
  const { content } = artifact;
  switch (artifactKind(artifact.name)) {
    case "html": {
      // The policy must come before any of the author's markup; keep a leading doctype first so the page stays out of quirks mode.
      const meta = `<meta http-equiv="Content-Security-Policy" content="${HTML_POLICY}">`;
      const doctype = /^\s*<!doctype[^>]*>/i.exec(content);
      return doctype ? doctype[0] + meta + content.slice(doctype[0].length) : meta + content;
    }
    case "svg": return shell(STATIC_POLICY, content);
    case "markdown": return shell(STATIC_POLICY, renderMarkdown(content));
    case "csv": return shell(STATIC_POLICY, csvTable(content));
    case "json": return shell(STATIC_POLICY, `<pre>${escapeHtml(prettyJson(content))}</pre>`);
    default: return shell(STATIC_POLICY, `<pre>${escapeHtml(content)}</pre>`);
  }
}

/** Card shown in the chat. Pressing it opens the preview. */
export function renderArtifactCard(artifact: StoredArtifact, open: (artifact: StoredArtifact) => void): HTMLButtonElement {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "artifact-card";
  card.title = `Preview ${artifact.name}`;
  const icon = document.createElement("span");
  icon.className = "artifact-card-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = artifactKind(artifact.name) === "html" ? "</>" : "▤";
  const text = document.createElement("span");
  text.className = "artifact-card-text";
  const name = document.createElement("strong");
  name.textContent = artifact.name;
  const kind = document.createElement("span");
  kind.textContent = `${KIND_LABEL[artifactKind(artifact.name)]} · Click to preview`;
  text.append(name, kind);
  card.append(icon, text);
  card.addEventListener("click", () => open(artifact));
  return card;
}
