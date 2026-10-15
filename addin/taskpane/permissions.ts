import type { ToolResult } from "../../protocol/src/index.js";

export type PermissionMode = "ask" | "auto";

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
