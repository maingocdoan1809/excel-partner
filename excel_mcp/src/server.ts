import { toolResultToMcp } from "../../shared/src/results.js";
import { describableTools, officeToolCatalog, unknownToolHint } from "../../shared/src/catalog.js";
import { DefinitionGate, GET_TOOL_DEFINITION, givenToolNames } from "../../shared/src/lazy-tools.js";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from "@modelcontextprotocol/sdk/types.js";
import { ZodError } from "zod";
import { DEFAULT_HOST, exposedToolNamesForHost, isKnownToolName, parseToolInputForHost, type HostId, type ToolError, type ToolGroups } from "../../protocol/src/index.js";
import type { ToolForwarder } from "./forwarder.js";

export const SERVER_INFO = { name: "excel-ai-addin", version: "0.1.0" } as const;

export interface ExcelMcpOptions {
  dynamicTools?: boolean;
  /** Expose opt-in tools such as execute_office_js. */
  allowOptInTools?: boolean;
  capabilities?: ToolGroups;
  /** The Office app the session belongs to (OFFICE_HOST); only that host's tools are listed and callable. Defaults to Excel. */
  host?: HostId;
}

function fail(error: ToolError): CallToolResult { return toolResultToMcp({ ok: false, error }); }

/**
 * Build the Excel MCP server. Every tool call is validated against the shared
 * protocol schemas and forwarded to the bridge, which routes it to the task
 * pane holding the live workbook.
 */
export function createExcelMcpServer(forwarder: ToolForwarder, options: ExcelMcpOptions = {}): Server {
  const host = options.host ?? DEFAULT_HOST;
  const allowOptIn = options.allowOptInTools ?? false;
  const groups = options.capabilities ?? {};
  const exposed = new Set<string>(exposedToolNamesForHost(host, allowOptIn, groups));
  // With lazy schemas, what the model has been told about each tool in this conversation (see DefinitionGate).
  const gate = groups.lazySchemas === true ? new DefinitionGate(describableTools(host, allowOptIn, groups)) : null;
  const server = new Server(SERVER_INFO, { capabilities: { tools: {} } });

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: options.dynamicTools && forwarder.listTools ? await forwarder.listTools() : officeToolCatalog(host, options.allowOptInTools ?? false, options.capabilities ?? {}),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request): Promise<CallToolResult> => {
    const name = request.params.name;
    if (options.dynamicTools && name.startsWith("ext_") && forwarder.listTools) {
      if (!(await forwarder.listTools()).some(tool => tool.name === name)) return fail({ code: "unknown_tool", message: `unknown tool '${name}'` });
      return toolResultToMcp(await forwarder.call(name, request.params.arguments ?? {}));
    }
    if (!isKnownToolName(name) || !exposed.has(name)) {
      return fail({ code: "unknown_tool", message: `unknown tool '${name}'.${unknownToolHint(host, allowOptIn, groups, name)}` });
    }

    const refusal = gate?.refuse(name, request.params.arguments) ?? gate?.unknownArguments(name, request.params.arguments);
    if (refusal) return fail({ code: "invalid_arguments", message: refusal });

    let parsed: unknown;
    try {
      parsed = parseToolInputForHost(host, name, request.params.arguments);
    } catch (error) {
      if (error instanceof ZodError) {
        return fail({
          code: "invalid_arguments",
          message: `invalid arguments for ${name}`,
          detail: error.issues,
        });
      }
      return fail({ code: "invalid_arguments", message: String(error) });
    }

    const result = await forwarder.call(name, parsed);
    if (gate && name === GET_TOOL_DEFINITION && result.ok) gate.noteGiven(givenToolNames(result.data));
    return toolResultToMcp(result);
  });

  return server;
}
