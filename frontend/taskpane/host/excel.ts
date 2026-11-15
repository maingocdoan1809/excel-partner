import { SHARED_WRITE_TOOLS, type ToolResult } from "../../../protocol/src/index.js";
import { WRITE_TOOLS, executeTool } from "../executor/index.js";
import { getSelectionLabel } from "../executor/read.js";
import { executeSharedTool } from "./shared.js";
import type { HostAdapter } from "./types.js";

async function readName(): Promise<string | null> {
  try {
    return await Excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
      return context.workbook.name || null;
    });
  } catch { return null; }
}

export const excelHost: HostAdapter = {
  id: "excel",
  app: "Excel",
  noun: "workbook",
  writeTools: new Set<string>([...WRITE_TOOLS, ...SHARED_WRITE_TOOLS]),
  execute: async (tool, args): Promise<ToolResult> => (await executeSharedTool("excel", tool, args)) ?? executeTool(tool, args),
  readName,
  readSelectionLabel: getSelectionLabel,
  hasSheets: true,
};
