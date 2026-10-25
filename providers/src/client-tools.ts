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

/** A provider may wrap MCP calls in its own tool protocol. Keep that translation
 * here rather than teaching API clients about CLI-specific calling conventions. */
export function clientToolInstructions(provider?: string): string {
  if (provider !== "opencode") return CLIENT_TOOL_INSTRUCTIONS;
  return [CLIENT_TOOL_INSTRUCTIONS,
    "OpenCode V2 Code Mode routing: use a targeted search for the exact excel tool needed, then reuse its discovered schema.",
    "Do not repeatedly dump the full tool catalog. If inventory is requested, list paths and short descriptions once, then look up full schemas only for the tools you will call.",
    "Inside the built-in execute tool, call await tools.excel.<remote_tool_name>(arguments) and return the remote result.",
    "Discover the remote input schema and pass exactly those arguments; do not assume a parameter name from another tool.",
    "If the client's remote tool is named execute, call tools.excel.execute(arguments), not the local execute sandbox with those arguments.",
    "Office code belongs in the code argument of that remote call. Never evaluate Office.js directly inside OpenCode's JavaScript sandbox.",
    "Older OpenCode versions expose prefixed MCP tools directly; use those instead of Code Mode when present.",
  ].join("\n");
}
