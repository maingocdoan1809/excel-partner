import type { StoredTool } from "./sessions.js";

const names: Record<string, string> = {
  read_range: "Reading cells", get_cell_ranges: "Reading cells", get_selection: "Reading selection",
  get_sheet_info: "Reading worksheet", get_objects: "Reading workbook objects", list_shapes: "Reading shapes",
  write_range: "Writing cells", set_cells: "Writing cells", set_formula: "Writing formulas",
  execute_office_js: "Running Excel code", read_range_image: "Viewing cells", manage_sheet: "Updating worksheet",
  format_range: "Formatting cells", modify_chart: "Updating chart", modify_table: "Updating table",
  get_sheets: "Reading sheets", search_data: "Searching cells", find_errors: "Checking formula errors",
  check_formula_consistency: "Checking formula consistency", find_replace: "Replacing text", get_range_as_csv: "Reading cells as CSV",
};
const tableTools = new Set(["read_range", "get_cell_ranges", "get_selection", "get_sheet_info", "get_sheets", "get_objects", "search_data"]);

function svg(paths: string[]): SVGSVGElement {
  const node = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  node.setAttribute("viewBox", "0 0 20 20");
  node.setAttribute("aria-hidden", "true");
  for (const data of paths) {
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", data); node.append(path);
  }
  return node;
}

export function toolChevron(): SVGSVGElement {
  return svg(["m7 5 5 5-5 5"]);
}

function toolIcon(tool: StoredTool): SVGSVGElement {
  const circle = "M17.5 10a7.5 7.5 0 1 1-15 0 7.5 7.5 0 0 1 15 0Z";
  if (tool.state === "error") return svg([circle, "m7.5 7.5 5 5m0-5-5 5"]);
  if (tool.state === "running") return svg([circle, "M10 5.5V10l3 1.5"]);
  return tableTools.has(tool.name)
    ? svg(["M2 4h16v12H2Z M2 8h16 M6 4v12"])
    : svg([circle, "m6.5 10 2.2 2.2 4.8-5"]);
}

function toolLabel(tool: StoredTool): string {
  // Use the tool's supplied purpose when present; never infer it from code.
  if (tool.name === "execute_office_js" && tool.input) {
    try { const input = JSON.parse(tool.input); if (typeof input.explanation === "string" && input.explanation.trim()) return input.explanation.trim().slice(0, 160); } catch {}
  }
  return names[tool.name] ?? tool.name.replace(/_/g, " ");
}

/** Bound local history size and do not render binary/image payloads as prose. */
export function formatToolValue(value: unknown): string {
  try {
    const text = JSON.stringify(value, (_key, item) => {
      if (item && typeof item === "object" && item.type === "image") return { type: "image", detail: "Image data omitted from text view" };
      if (typeof item === "string" && item.length > 8000) return `${item.slice(0, 8000)}… [truncated; ${item.length} characters]`;
      return item;
    }, 2) ?? "null";
    return text.length > 16000 ? `${text.slice(0, 16000)}\n… [Display truncated]` : text;
  } catch { return "This value could not be displayed as JSON."; }
}

function jsonCode(text: string): HTMLElement {
  const code = document.createElement("code");
  const pattern = /"(?:\\.|[^"\\])*"\s*:|"(?:\\.|[^"\\])*"|\b(?:true|false|null)\b|-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b/g;
  let offset = 0;
  for (const match of text.matchAll(pattern)) {
    code.append(document.createTextNode(text.slice(offset, match.index)));
    const token = document.createElement("span");
    token.className = match[0].endsWith(":") ? "json-key" : match[0].startsWith('"') ? "json-string" : "json-literal";
    token.textContent = match[0];
    code.append(token); offset = match.index! + match[0].length;
  }
  code.append(document.createTextNode(text.slice(offset)));
  return code;
}

export function renderToolCard(tool: StoredTool): HTMLDetailsElement {
  const card = document.createElement("details");
  card.className = "tool-call";
  const summary = document.createElement("summary");
  const icon = document.createElement("span");
  icon.className = "tool-call-icon";
  icon.setAttribute("aria-hidden", "true");
  const name = document.createElement("span");
  name.className = "tool-call-name";
  name.textContent = toolLabel(tool);
  name.title = tool.name;
  const state = document.createElement("span"); state.className = "tool-call-state";
  const chevron = document.createElement("span"); chevron.className = "tool-call-chevron"; chevron.append(toolChevron());
  summary.append(icon, name, state, chevron);
  const body = document.createElement("div"); body.className = "tool-call-body";
  card.append(summary, body);
  paintTool(card, tool);
  return card;
}

export function paintTool(card: HTMLElement, tool: StoredTool): void {
  card.dataset.state = tool.state;
  card.querySelector(".tool-call-icon")!.replaceChildren(toolIcon(tool));
  card.querySelector(".tool-call-state")!.textContent = tool.state === "running" ? "Running…" : tool.state === "error" ? "Failed" : "";
  const body = card.querySelector(".tool-call-body")!;
  body.replaceChildren();
  for (const [label, value] of [["Arguments", tool.input || tool.args], [tool.error ? "Error" : "Result", tool.result ?? tool.error]] as const) {
    if (!value) continue;
    const section = document.createElement("section"); section.className = "tool-call-section";
    const heading = document.createElement("div"); heading.className = "tool-call-label"; heading.textContent = label;
    const pre = document.createElement("pre"); pre.append(jsonCode(value));
    section.append(heading, pre); body.append(section);
  }
  if (!tool.result && !tool.error) {
    const note = document.createElement("p"); note.className = "tool-call-unavailable";
    note.textContent = tool.state === "running" ? "Waiting for the tool result…" : "This provider did not expose a result body.";
    body.append(note);
  }
}
