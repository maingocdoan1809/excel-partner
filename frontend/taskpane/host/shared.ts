import type { CreateArtifactInput, ExecuteOfficeJsInput, HostId, ResolvedInsertImageInput, ToolResult } from "../../../protocol/src/index.js";
import { toToolError } from "../executor/common.js";
import { insertImage as insertExcelImage } from "../executor/images.js";
import { executeOfficeJs } from "../executor/office-js.js";
import { insertImage as insertPptImage } from "../executor/powerpoint/images.js";
import { insertImage as insertWordImage } from "../executor/word/images.js";

/** insert_image per host: how a picture gets into a workbook, a document and a slide. */
const IMAGE_INSERTERS: Record<HostId, (input: ResolvedInsertImageInput) => Promise<unknown>> = {
  excel: insertExcelImage,
  word: insertWordImage,
  powerpoint: insertPptImage,
};

/** Tools every host has and that do not depend on what the document holds. Null when `tool` is not one of them. */
export async function executeSharedTool(host: HostId, tool: string, args: unknown): Promise<ToolResult | null> {
  try {
    if (tool === "execute_office_js") return { ok: true, data: await executeOfficeJs(args as ExecuteOfficeJsInput, host) };
    if (tool === "insert_image") {
      // The bridge read the picture and added it to the call; a call without it did not come from the bridge.
      const input = args as ResolvedInsertImageInput;
      if (typeof input?.imageBase64 !== "string" || !input.imageBase64) return { ok: false, error: { code: "invalid_arguments", message: "insert_image arrived without a picture" } };
      return { ok: true, data: await IMAGE_INSERTERS[host](input) };
    }
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
