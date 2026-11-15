import type { HostId } from "../../protocol/src/index.js";
import type { JsonSchema, ToolDefinition } from "./tool-definitions.js";

/**
 * Lazy tool schemas. Every tool is still registered with the provider (so it can be called by its own name, shows its own
 * approval card and keeps its read/write annotations), but all except a few carry only a one-line summary and an open input
 * schema. The model asks `get_tool_definition` for the real schema before first use; a wrong call gets the schema back in
 * its error, so skipping that step costs one retry, not a failure.
 */

export const GET_TOOL_DEFINITION = "get_tool_definition";

/** Tools that keep their full schema in lazy mode: what nearly every request on that host starts with. */
export const HOT_TOOLS: Readonly<Record<HostId, readonly string[]>> = {
  excel: ["get_selection", "get_sheets", "read_range", "write_range", "search_data", "set_formula", "format_range"],
  word: ["get_selection", "get_doc_outline", "read_doc_section", "edit_doc_text", "edit_doc_paragraphs"],
  powerpoint: ["get_selection", "get_presentation_metadata", "list_slide_shapes", "read_slide_text", "edit_slide_text"],
};

/** The open input schema of a stub: anything is accepted by the provider and checked by the real schema on our side. */
const OPEN_SCHEMA: JsonSchema = { type: "object", additionalProperties: true };

/** One-line summaries for tools whose description does not open with a sentence that stands alone. */
export const TOOL_SUMMARIES: Readonly<Record<string, string>> = {
  modify_structure: "WRITE. Insert, delete, hide, unhide, group or ungroup rows/columns; freeze panes; collapse or expand outline levels.",
  conditional_format: "WRITE. Add, list or clear conditional formatting: value, text, formula, colour-scale, data-bar, icon-set, top/bottom and duplicate rules.",
  data_validation: "WRITE. Set, get or clear data validation: dropdown lists, number/date/time/text-length limits and custom formulas.",
  manage_content_controls: "WRITE (operation list only reads). List, fill, insert or delete content controls by tag or title.",
  manage_doc_pictures: "WRITE (operation list only reads). List the inline pictures; resize one or set its alt text; delete one.",
  read_connected_file: "READ ONLY. Run one read tool of that file's own application in ANOTHER open file (see list_connected_files).",
  send_to_connected_file: "Send a request to the assistant of ANOTHER open file, which does the work there and replies.",
  trace_formula: "READ. Follow a cell's formula: what it reads (precedents) and what reads it (dependents), up to 3 steps.",
  batch: "WRITE. Run up to 20 listed tools in order in one call, each with the arguments it takes alone; stops at the first failure.",
  analyze_range: "READ. Profile a table's columns, or filter, group and aggregate its rows (totals, counts, rankings) without reading them.",
  get_selection: "READ. What is selected now: cells, text or slide shapes, depending on the open application.",
  insert_image: "WRITE. Put a picture (web address, file or base64) into the open file where it is selected.",
  update_skill: "WRITE. Update a text file of a user-installed skill (SKILL.md, references or scripts).",
  search_data: "READ. Find cells containing text across the visible sheets (or one sheet/range), optionally as a regex or inside formulas.",
  validate_flowchart: "READ. Check the connected shapes of a sheet: disconnected nodes, unattached connectors, overlaps and crossings.",
  page_layout: "WRITE. Print area, orientation, paper size, scaling, repeating titles, margins, header/footer, print options and page breaks.",
  create_artifact: "DISPLAY. Show the user a file (.html, .svg, .md, .csv, .json, .txt) in a preview next to the chat.",
  list_test_sessions: "READ. The signed-in sessions the user saved for testing web systems, with the sites each may reach.",
  open_test_browser: "WRITE. Start the Playwright browser, signed in with a saved session, limited to the allowed sites.",
  read_site_knowledge: "READ. Notes saved from earlier sessions about one web system: list them, or read one by name.",
  save_site_knowledge: "WRITE. Keep a short note about a web system (login, navigation, how-to, quirk, safe test data) for later sessions.",
  read_doc_section: "READ. Read the section under a heading, or `count` paragraphs from an index, with style, text and list/table flags.",
  search_doc_text: "READ. Find text in the body: total matches and, for the first ones, paragraph index, occurrence number and context.",
  edit_doc_list: "WRITE. Insert a NEW bulleted or numbered list (nesting allowed) after a paragraph or at the end.",
  propose_doc_edits: "WRITE (staged). Show a review card of text edits; only the ones the person accepts are applied.",
  apply_style: "WRITE. Set paragraph style, font, alignment or spacing on the selection or on a run of paragraphs.",
  edit_table: "WRITE (operation read only reads). Work on a Word table by index: read, set a cell, add/delete rows and columns, merge, style.",
  list_slide_shapes: "READ. The shapes of a slide: ref, name, type, placeholder kind, position, size, stacking and a text excerpt.",
  read_slide_notes: "READ. The speaker notes of the given slides (or all) as plain text.",
  edit_slide_xml: "WRITE. Change a slide with code on its raw OOXML (formatting-preserving text, charts, transitions): the escape hatch of the typed tools.",
  insert_slide_element: "WRITE. Add one textbox, shape, line or table to a slide at a position and size in points.",
  verify_slides: "READ. Geometry checks as estimates: shapes off the slide, overlaps, empty placeholders, text overflow, tiny fonts, low contrast.",
  align_shapes: "WRITE. Align shapes and/or distribute them evenly, relative to each other or to the slide.",
  get_presentation_metadata: "READ. Start here: file name, slide count and size, every slide's id, position, title and layout, masters, API level.",
  edit_slide_table: "WRITE (operation read only reads). Edit a slide table: set a cell, add or delete rows and columns.",
};

export const MAX_SUMMARY_CHARS = 160;
const ABBREVIATIONS = /(?:e\.g|i\.e|etc|vs|approx|no)\.$/i;
/** Descriptions open with what the tool does to the file: `READ.`, `WRITE.`, `WRITE (operation list only reads).`, `EXECUTE.` */
const ACCESS_TAG = /^((?:READ|WRITE|EXECUTE|DISPLAY|CHECK)\b[^.]{0,40}\.)\s+/;

/** The first sentence of a description after its access tag, with the tag kept in front; null when it is not a short sentence. */
export function firstSentence(description: string): string | null {
  const text = description.replace(/\s+/g, " ").trim();
  const tag = ACCESS_TAG.exec(text);
  const prefix = tag ? `${tag[1]} ` : "";
  const body = tag ? text.slice(tag[0].length) : text;
  const room = MAX_SUMMARY_CHARS - prefix.length;
  for (let i = 0; i < body.length && i < room + 40; i++) {
    const ch = body[i]!;
    if (ch !== "." && ch !== "!" && ch !== "?") continue;
    const next = body[i + 1];
    if (next !== undefined && next !== " ") continue; // "1.19", "A1.B2", "path.ext"
    if (ch === "." && ABBREVIATIONS.test(body.slice(0, i + 1))) continue;
    const sentence = body.slice(0, i + 1);
    return sentence.length <= room ? prefix + sentence : null;
  }
  return body.length <= room ? prefix + body : null;
}

/** What a stub says about the tool: a hand-written summary where one exists, else the description's first sentence. */
export function summaryOf(tool: Pick<Described, "name" | "description">): string {
  return TOOL_SUMMARIES[tool.name] ?? firstSentence(tool.description) ?? tool.description.slice(0, MAX_SUMMARY_CHARS - 1).trimEnd() + "…";
}

/** The definition a lazy catalog offers in place of the real one. Annotations stay: providers use them for approval. */
export function stubOf(tool: ToolDefinition): ToolDefinition {
  return { ...tool, description: `${summaryOf(tool)} Call ${GET_TOOL_DEFINITION} before first use.`, inputSchema: OPEN_SCHEMA };
}

export function isHotTool(host: HostId, name: string): boolean {
  return name === GET_TOOL_DEFINITION || HOT_TOOLS[host].includes(name);
}

// --- Finding a tool from a name the model may have got slightly wrong -----------------------------------------------------

/** `Read Range`, `read-range`, `mcp__excel__read_range` and `ext_x__read_range` all mean `read_range`. */
export function normalizeToolName(name: string): string {
  const last = name.split("__").pop() ?? name;
  return last.trim().replace(/([a-z0-9])([A-Z])/g, "$1_$2").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function editDistance(a: string, b: string, limit: number): number {
  if (Math.abs(a.length - b.length) > limit) return limit + 1;
  let previous = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) row[j] = Math.min(previous[j]! + 1, row[j - 1]! + 1, previous[j - 1]! + (a[i - 1] === b[j - 1] ? 0 : 1));
    previous = row;
  }
  return previous[b.length]!;
}

/** Up to `max` tool names that the (misspelt, shortened or re-cased) `input` plausibly means, best first. Empty when none comes close. */
export function suggestToolNames(input: string, names: readonly string[], max = 3): string[] {
  const wanted = normalizeToolName(input);
  if (!wanted) return [];
  const scored: Array<{ name: string; score: number }> = [];
  for (const name of names) {
    const candidate = normalizeToolName(name);
    let score: number | null = null;
    if (candidate === wanted) score = 0;
    else if (candidate.startsWith(wanted) || wanted.startsWith(candidate)) score = 1;
    else if (candidate.includes(wanted) || wanted.includes(candidate)) score = 2;
    else {
      const limit = Math.max(2, Math.floor(wanted.length / 4));
      const distance = editDistance(wanted, candidate, limit);
      if (distance <= limit) score = 2 + distance;
      else if (wanted.split("_").some(part => part.length > 3 && candidate.split("_").includes(part))) score = 20; // shares a whole word
    }
    if (score !== null) scored.push({ name, score });
  }
  return scored.sort((a, b) => a.score - b.score || a.name.length - b.name.length || a.name.localeCompare(b.name)).slice(0, max).map(entry => entry.name);
}

type Described = Pick<ToolDefinition, "name" | "description" | "inputSchema">;

/** Most definitions one answer carries, so a request for "everything" cannot put the whole catalog back into the conversation. */
export const MAX_DEFINITIONS = 12;

export interface ToolLookup {
  /** Full definitions of the tools that were found, in the order asked. */
  definitions: Described[];
  /** What to do next, said plainly: given whenever something asked for was not found or not all of it fit. */
  note?: string;
  /** Names that matched nothing exactly, each with the closest tools. */
  unmatched: Array<{ input: string; suggestions: string[] }>;
  /** Every tool and what it does: given only when nothing asked for was found (or nothing was asked). */
  tools?: Array<{ name: string; summary: string }>;
}

/** The names a `get_tool_definition` call asked for: `names`, or the single `name` a model may send instead. */
export function requestedToolNames(args: unknown): string[] {
  const record = args && typeof args === "object" ? args as Record<string, unknown> : {};
  if (Array.isArray(record.names)) return record.names.filter((name): name is string => typeof name === "string");
  return typeof record.name === "string" ? [record.name] : [];
}

/**
 * Answer `get_tool_definition`. A name that is only a close match is still taken when one tool is clearly meant (same name
 * once normalised); otherwise it is reported with suggestions. When nothing matches, or nothing was asked, the whole list
 * comes back so the model can pick.
 */
export function lookupTools(requested: readonly string[], full: readonly Described[]): ToolLookup {
  const byName = new Map(full.map(tool => [tool.name, tool] as const));
  const byNormalized = new Map(full.map(tool => [normalizeToolName(tool.name), tool] as const));
  const definitions: Described[] = [];
  const unmatched: ToolLookup["unmatched"] = [];
  let leftOut = 0;
  const names = full.map(tool => tool.name);
  for (const input of requested) {
    const tool = byName.get(input) ?? byNormalized.get(normalizeToolName(input));
    if (!tool) { unmatched.push({ input, suggestions: suggestToolNames(input, names) }); continue; }
    if (definitions.some(found => found.name === tool.name)) continue;
    if (definitions.length >= MAX_DEFINITIONS) leftOut++;
    else definitions.push({ name: tool.name, description: tool.description, inputSchema: tool.inputSchema });
  }
  const result: ToolLookup = { definitions, unmatched };
  if (!definitions.length) {
    result.tools = full.map(tool => ({ name: tool.name, summary: summaryOf(tool) }));
    result.note = requested.length ? `None of these names is a tool here. Every tool is listed below; call ${GET_TOOL_DEFINITION} again with names from the list.` : `Every tool is listed below; call ${GET_TOOL_DEFINITION} with the names you need.`;
  } else if (unmatched.length) result.note = `Not found: ${unmatched.map(item => item.input).join(", ")}. See suggestions for the closest tools.`;
  if (leftOut) result.note = `${result.note ? result.note + " " : ""}${leftOut} more tool(s) were left out to keep this short: ask for them in another call.`;
  return result;
}
