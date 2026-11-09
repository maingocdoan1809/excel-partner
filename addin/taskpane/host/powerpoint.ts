import { PPT_WRITE_TOOLS, type ToolResult } from "../../../protocol/src/index.js";
import { toToolError } from "../executor/common.js";
import { executePptTool, getPptSelection, getPptSelectionLabel, readPptPresentationName } from "../executor/powerpoint/index.js";
import { readOpenFile } from "./file.js";
import { executeSharedTool } from "./shared.js";
import type { HostAdapter } from "./types.js";

export const powerpointHost: HostAdapter = {
  id: "powerpoint",
  app: "PowerPoint",
  noun: "presentation",
  extension: ".pptx",
  writeTools: new Set<string>([...PPT_WRITE_TOOLS, "execute_office_js"]),
  execute: async (tool, args): Promise<ToolResult> => {
    const shared = await executeSharedTool("powerpoint", tool, args);
    if (shared) return shared;
    if (tool === "get_selection") {
      try { return { ok: true, data: await getPptSelection() }; } catch (error) { return { ok: false, error: toToolError(error) }; }
    }
    return executePptTool(tool, args);
  },
  readName: readPptPresentationName,
  readSelectionLabel: getPptSelectionLabel,
  readFile: () => readOpenFile("powerpoint"),
  hasSheets: false,
};
