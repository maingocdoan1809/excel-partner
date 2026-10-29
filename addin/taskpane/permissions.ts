import type { ToolResult } from "../../protocol/src/index.js";
import type { RunSkillScriptInput } from "../../protocol/src/index.js";

export type PermissionMode = "ask" | "auto";

export async function authorizeSkillScript(
  mode: PermissionMode,
  input: RunSkillScriptInput,
  ask: (detail: string, target: string) => Promise<boolean>,
): Promise<ToolResult> {
  const approved = mode === "auto" || await ask(
    "Run this skill script on your computer? It can read, change, or delete local files with your user permissions. " +
    `Time limit: ${input.timeoutMs / 1000} seconds.`,
    `${input.name}/${input.file}\nArguments: ${JSON.stringify(input.args)}${input.stdin === undefined ? "" : `\nStandard input: ${input.stdin.slice(0, 4000)}`}`,
  );
  return { ok: true, data: { approved } };
}

const OVERWRITE_TOOLS = new Set(["write_range", "set_formula", "set_cells", "copy_to", "merge_range"]);

function withPermissionFlag(tool: string, args: unknown, enabled: boolean): unknown {
  if (!args || typeof args !== "object" || Array.isArray(args)) return args;
  if (OVERWRITE_TOOLS.has(tool)) return { ...args, allowOverwrite: enabled };
  if (tool === "manage_sheet" && (args as { operation?: unknown }).operation === "delete") {
    return { ...args, confirmDelete: enabled };
  }
  return args;
}

/** Execute once with guards, then ask the user and retry only the refused operation. */
export async function executeWithPermission(
  tool: string,
  args: unknown,
  mode: PermissionMode,
  execute: (tool: string, args: unknown) => Promise<ToolResult>,
  ask: (message: string) => Promise<boolean>,
): Promise<ToolResult> {
  const prepared = withPermissionFlag(tool, args, mode === "auto");
  const result = await execute(tool, prepared);
  if (mode === "auto" || result.ok) return result;
  const flag = result.error.code === "would_overwrite" && OVERWRITE_TOOLS.has(tool)
    ? "allowOverwrite"
    : result.error.code === "confirmation_required" && tool === "manage_sheet"
      ? "confirmDelete"
      : null;
  if (!flag) return result;
  const approved = await ask(result.error.message.split(" Ask the user")[0]!);
  if (!approved) return { ok: false, error: { code: "cancelled", message: "The user declined this workbook change." } };
  return execute(tool, { ...(prepared as Record<string, unknown>), [flag]: true });
}

export interface TestBrowserApproval {
  origins: string[];
  visible: boolean;
  /** Name of the saved sign-in the browser starts with. */
  session?: string;
}

/** Wording of the approval dialog for the test browser. It is asked in Auto mode too. */
export function describeTestBrowserApproval(args: TestBrowserApproval): { title: string; detail: string; confirm: string; target: string } {
  const how = args.session
    ? `already signed in with your saved session "${args.session}" (the assistant never sees the cookies or the password)`
    : "signed out, in a window where you sign in yourself (the password is never shared with the assistant)";
  return {
    title: "Open a test browser?",
    detail: `The assistant will drive a separate browser, ${how}, ${args.visible ? "in a window you can see" : "hidden (you watch it in the task pane)"}. `
      + "It can open only the sites below; every other address, including IP addresses, is blocked.",
    confirm: "Open browser", target: args.origins.join("\n"),
  };
}

export interface SiteKnowledgeApproval {
  action: "save" | "delete";
  origin: string;
  name: string;
  kind?: string;
  summary?: string;
  content?: string;
  replaces?: boolean;
}

/** Wording of the dialog for changing the saved notes about a system. Notes steer later sessions, so Ask mode shows them in full. */
export function describeSiteKnowledgeApproval(args: SiteKnowledgeApproval): { title: string; detail: string; confirm: string; target: string } {
  if (args.action === "delete") {
    return { title: "Delete a saved note?", detail: `Delete the note "${args.name}" about ${args.origin}. Future sessions will no longer see it.`, confirm: "Delete note", target: "" };
  }
  return {
    title: args.replaces ? "Replace a saved note?" : "Save a note about this system?",
    detail: `The assistant wants to keep a ${args.kind ?? "note"} note "${args.name}" about ${args.origin} on this computer, so later sessions start with it: ${args.summary ?? ""} `
      + "It is plain text you can read or delete in the knowledge folder (excel-partner knowledge). It must not hold passwords or tokens.",
    confirm: args.replaces ? "Replace note" : "Save note", target: (args.content ?? "").slice(0, 8000),
  };
}
