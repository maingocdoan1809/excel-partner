import { clientToolInstructions } from "../../providers/src/client-tools.js";
import { OFFICE_TOOL_GUIDANCE } from "./office-tool-guidance.js";

/** Both the model session and MCP initialization receive the same API contract.
 * Some CLI transports do not surface MCP server instructions to the model. */
export function gatewayToolInstructions(provider?: string, officeJs = false): string {
  return [clientToolInstructions(provider), ...(officeJs ? [OFFICE_TOOL_GUIDANCE] : [])].join("\n");
}
