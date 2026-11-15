import type { ToolError, ToolResult } from "../../protocol/src/index.js";

/** Carries a failed tool result out of code that is not itself a tool (a file lookup, a permission question) to the caller that returns it. */
export class ToolResultError extends Error {
  constructor(readonly result: ToolResult) {
    super(result.ok ? "tool result" : result.error.message);
  }
}

/** A failed tool result. `detail` is left out entirely when absent so the JSON stays small. */
export function toolError(code: ToolError["code"], message: string, detail?: unknown): ToolResult {
  return { ok: false, error: detail === undefined ? { code, message } : { code, message, detail } };
}
