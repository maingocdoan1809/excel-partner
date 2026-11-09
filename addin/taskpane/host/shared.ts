import type { CreateArtifactInput, ExecuteOfficeJsInput, HostId, ToolResult } from "../../../protocol/src/index.js";
import { toToolError } from "../executor/common.js";
import { executeOfficeJs } from "../executor/office-js.js";

/** Tools every host has and that do not depend on what the document holds. Null when `tool` is not one of them. */
export async function executeSharedTool(host: HostId, tool: string, args: unknown): Promise<ToolResult | null> {
  try {
    if (tool === "execute_office_js") return { ok: true, data: await executeOfficeJs(args as ExecuteOfficeJsInput, host) };
    if (tool === "create_artifact") {
      // The task pane draws the preview card from the call itself; the tool only acknowledges it.
      const { filename, content } = args as CreateArtifactInput;
      return { ok: true, data: { filename, characters: content.length, shown: true } };
    }
  } catch (error) {
    return { ok: false, error: toToolError(error) };
  }
  return null;
}
