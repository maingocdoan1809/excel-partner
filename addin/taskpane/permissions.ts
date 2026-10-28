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

export interface QaApproval {
  action: "save" | "run" | "run_draft" | "open_session";
  /** Ask even in Auto mode (unreviewed draft scripts). */
  forceAsk?: boolean;
  kind?: "test" | "page";
  id?: string;
  file?: string;
  code?: string;
  environment?: string;
  baseUrl?: string;
  scripts?: Array<{ tcId: string; title: string; role: string; status: string }>;
  profile?: string;
  origins?: string[];
}

/** Wording of the approval dialog for QA harness actions. */
export function describeQaApproval(args: QaApproval): { title: string; detail: string; confirm: string; target: string } {
  const list = (args.scripts ?? []).map(item => `${item.tcId}  ${item.title}  [${item.role}, ${item.status}]`).join("\n");
  switch (args.action) {
    case "save":
      return {
        title: "Save test script?",
        detail: `Save ${args.kind === "page" ? "page object" : "test"} ${args.id ?? ""} as a draft in your QA project (${args.file ?? ""}). `
          + (args.kind === "page" ? "Tests that use it will need to be approved again." : `It will not run unattended until you approve it: excel-partner qa approve ${args.id ?? ""}.`),
        confirm: "Save draft", target: (args.code ?? "").slice(0, 8000),
      };
    case "run_draft":
      return {
        title: "Run UNAPPROVED test scripts?",
        detail: `These scripts are drafts that nobody has approved. A visible browser will run them against ${args.baseUrl ?? "the staging site"} (${args.environment ?? ""}) with the saved session(s) for: ${[...new Set((args.scripts ?? []).map(item => item.role))].join(", ")}. The browser can reach only that site. Read the scripts first if unsure.`,
        confirm: "Run drafts", target: list,
      };
    case "run":
      return {
        title: "Run automated tests?",
        detail: `Run ${(args.scripts ?? []).length} approved script(s) against ${args.baseUrl ?? "the staging site"} (${args.environment ?? ""}). The browser can reach only that site.`,
        confirm: "Run tests", target: list,
      };
    case "open_session":
      return {
        title: "Open a signed-in test browser?",
        detail: `The assistant will browse ${args.environment ?? "the staging site"} as '${args.profile ?? ""}' in a separate browser limited to: ${(args.origins ?? []).join(", ")}. It uses the saved session; the password is never shared.`,
        confirm: "Open browser", target: "",
      };
  }
}
