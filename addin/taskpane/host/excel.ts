import type { ToolResult } from "../../../protocol/src/index.js";
import { WRITE_TOOLS, executeTool } from "../executor/index.js";
import { getSelectionLabel } from "../executor/read.js";
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
  writeTools: WRITE_TOOLS,
  execute: (tool, args): Promise<ToolResult> => executeTool(tool, args),
  readName,
  readSelectionLabel: getSelectionLabel,
  hasSheets: true,
};
