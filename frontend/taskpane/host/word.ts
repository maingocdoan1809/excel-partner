import { SHARED_WRITE_TOOLS, WORD_WRITE_TOOLS, type ToolResult } from "../../../protocol/src/index.js";
import { executeWordTool, getWordSelectionLabel, readWordDocumentName } from "../executor/word/index.js";
import { executeSharedTool } from "./shared.js";
import type { HostAdapter } from "./types.js";

export const wordHost: HostAdapter = {
  id: "word",
  app: "Word",
  noun: "document",
  writeTools: new Set<string>([...WORD_WRITE_TOOLS, ...SHARED_WRITE_TOOLS]),
  execute: async (tool, args, signal): Promise<ToolResult> =>
    (await executeSharedTool("word", tool, args)) ?? executeWordTool(tool, args, signal),
  readName: readWordDocumentName,
  readSelectionLabel: getWordSelectionLabel,
  hasSheets: false,
};
