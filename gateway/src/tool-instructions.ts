import type { HostId } from "../../protocol/src/index.js";
import { clientToolInstructions } from "../../providers/src/client-tools.js";
import { OFFICE_TOOL_GUIDANCE } from "./office-tool-guidance.js";

/** Both the model session and MCP initialization receive the same API contract.
 * Some CLI transports do not surface MCP server instructions to the model. */
export function gatewayToolInstructions(providerNotes: string | undefined, officeJs = false, host?: HostId): string {
  return [clientToolInstructions(providerNotes), ...(host && host !== "excel" ? [`The API client is working in ${host === "word" ? "Word" : "PowerPoint"}. Use only the tools and schemas it supplied; do not assume Office Partner tool names. Office.js code must use ${host === "word" ? "Word" : "PowerPoint"} APIs, with context provided by the client tool contract. Load properties and sync before reading; verify after edits.`] : officeJs ? ["Apply the following Excel-specific API guidance only when the client is working in Excel. For Word or PowerPoint use that host’s APIs and the client’s tool contract.", OFFICE_TOOL_GUIDANCE] : [])].join("\n");
}
