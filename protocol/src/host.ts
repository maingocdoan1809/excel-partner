import { z } from "zod";
import {
  EXCEL_TOOL_NAMES,
  OPT_IN_TOOLS,
  parseToolInput,
  toolInGroups,
  type ExcelToolName,
  type ToolGroups,
  type ToolName,
} from "./registry.js";
import { PPT_TOOL_NAMES } from "./tools-ppt.js";
import { WORD_TOOL_NAMES } from "./tools-word.js";

/** The Office application a task pane runs in. Standalone browser panes count as "excel". */
export const HostIdSchema = z.enum(["excel", "word", "powerpoint"]);
export type HostId = z.infer<typeof HostIdSchema>;
export const HOST_IDS: readonly HostId[] = HostIdSchema.options;
export const DEFAULT_HOST: HostId = "excel";

export function isHostId(value: unknown): value is HostId {
  return typeof value === "string" && (HOST_IDS as readonly string[]).includes(value);
}

/**
 * Tools that exist for every host: the pane's selection / escape-hatch tools, and tools the bridge answers itself
 * without touching the document (skills, Python, the browser, attachments, preview cards).
 */
export const HOST_AGNOSTIC_TOOLS: readonly ExcelToolName[] = [
  "get_selection",
  "execute_office_js",
  "create_artifact",
  "list_skills",
  "read_skill",
  "save_skill",
  "update_skill",
  "run_skill_script",
  "run_python",
  "playwright_browser",
  "list_reference_files",
  "read_workbook_file",
  "read_document_file",
  "read_attached_image",
  "read_attached_file",
  "list_connected_files",
  "read_connected_file",
  "send_to_connected_file",
];

/** Tools only Word has. */
export const WORD_TOOLS: readonly string[] = WORD_TOOL_NAMES;

/** Tools only PowerPoint has. */
export const PPT_TOOLS: readonly string[] = PPT_TOOL_NAMES;

/**
 * Which hosts each tool is available in. Every tool the Excel catalog already has is Excel-only unless it is
 * host-agnostic; key order starts with the Excel catalog order so `toolsForHost("excel")` equals the old list.
 */
export const TOOL_HOSTS: Readonly<Record<string, readonly HostId[]>> = {
  ...Object.fromEntries(EXCEL_TOOL_NAMES.map((name): [string, readonly HostId[]] => [
    name,
    (HOST_AGNOSTIC_TOOLS as readonly string[]).includes(name) ? HOST_IDS : ["excel"],
  ])),
  ...Object.fromEntries(WORD_TOOLS.map((name): [string, readonly HostId[]] => [name, ["word"]])),
  ...Object.fromEntries(PPT_TOOLS.map((name): [string, readonly HostId[]] => [name, ["powerpoint"]])),
};

export interface ToolExposureOptions extends ToolGroups {
  /** Whether opt-in tools such as `execute_office_js` are enabled (BRIDGE_ALLOW_OFFICE_JS). */
  allowOptIn: boolean;
}

/** Tool names a session of `host` exposes. The single source for MCP `ListTools`, the bridge gate and provider fallbacks. */
export function toolsForHost(host: HostId, opts: ToolExposureOptions): string[] {
  return Object.keys(TOOL_HOSTS).filter((name) => TOOL_HOSTS[name]!.includes(host)
    && (opts.allowOptIn || !(OPT_IN_TOOLS as readonly string[]).includes(name))
    && toolInGroups(name, opts));
}

/** Same as `exposedToolNames` for a host: `exposedToolNamesForHost("excel", a, c)` equals `exposedToolNames(a, c)`. */
export function exposedToolNamesForHost(host: HostId, allowOptIn: boolean, capabilities: ToolGroups = {}): string[] {
  return toolsForHost(host, { allowOptIn, ...capabilities });
}

/** Whether a tool belongs to the host's catalog at all, regardless of whether it is currently enabled. */
export function isToolNameForHost(host: HostId, name: string): boolean {
  return Object.prototype.hasOwnProperty.call(TOOL_HOSTS, name) && TOOL_HOSTS[name]!.includes(host);
}

/**
 * Parse and validate tool arguments for a host. Throws a ZodError for bad arguments and a plain Error when the tool is
 * not part of the host's catalog. Host-specific schemas (Word, PowerPoint) are registered here as they are added.
 */
export function parseToolInputForHost(host: HostId, tool: string, args: unknown): unknown {
  if (!isToolNameForHost(host, tool)) throw new Error(`tool '${tool}' is not available in ${host}`);
  return parseToolInput(tool as ToolName, args);
}

/** The host an MCP process / provider session was started for (`OFFICE_HOST`). Unset means Excel; anything else invalid throws. */
export function hostFromEnv(env: Record<string, string | undefined>): HostId {
  const value = env.OFFICE_HOST;
  if (value === undefined || value === "") return DEFAULT_HOST;
  const parsed = HostIdSchema.safeParse(value);
  if (!parsed.success) throw new Error(`OFFICE_HOST must be one of ${HOST_IDS.join(", ")}, got '${value}'`);
  return parsed.data;
}
