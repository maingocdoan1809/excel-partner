/** Shared contract for caller-owned tools, independent of provider transport. */
export const CLIENT_TOOL_CONTRACT_REVISION = "client-tools-v5";
export const CLIENT_TOOL_INSTRUCTIONS = [
  "The tools on the excel MCP server belong to the connected client and run in that client's application.",
  "Resolve tool names in the client's instructions to these remote MCP tools, never to similarly named provider built-ins.",
  "In particular, a remote execute or execute_office_js tool is not the provider's local JavaScript execute sandbox.",
  "Forward Office.js/workbook code as arguments to the remote tool; the local sandbox has no Excel workbook or Office context.",
  "Defining a function, returning null, or successfully running local JavaScript does not confirm any workbook change.",
  "Inspect the actual result shape before using it. A client may return JSON text or an envelope such as {success:true,result:{...}}; application fields then live under result, not at the top level.",
  "If MCP content contains JSON text, parse that text first. Treat success:false and isError:true as failures, not as empty application data.",
  "Never replace a missing or unparsed field with [] or false to decide that a resource does not exist. A failed read is not evidence of absence.",
  "For create-if-missing workbook operations, check the resource and conditionally create it in the same remote Office.js call, then read back its state. Do not repeat a write merely to return its result.",
  "Only report a remote action as successful after the client tool result confirms it. On errors, correct the call or report the failure.",
].join("\n");

/** The contract every client-tool transport shares, plus the notes a provider adds for its own way of calling MCP tools. */
export function clientToolInstructions(notes?: string): string {
  return notes ? [CLIENT_TOOL_INSTRUCTIONS, notes].join("\n") : CLIENT_TOOL_INSTRUCTIONS;
}
