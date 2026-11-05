import type { ToolError, ToolResult } from "../../protocol/src/index.js";

/** A failed tool result. `detail` is left out entirely when absent so the JSON stays small. */
export function toolError(code: ToolError["code"], message: string, detail?: unknown): ToolResult {
  return { ok: false, error: detail === undefined ? { code, message } : { code, message, detail } };
}
