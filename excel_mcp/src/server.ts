import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from "@modelcontextprotocol/sdk/types.js";
import { ZodError } from "zod";
import { exposedToolNames, isExcelToolName, parseToolInput, type ToolError } from "../../protocol/src/index.js";
import type { ToolForwarder } from "./forwarder.js";
import { ALL_TOOL_DEFINITIONS } from "./tool-definitions.js";

export const SERVER_INFO = { name: "excel-ai-addin", version: "0.1.0" } as const;

/** Tool results may carry one rendered image (`image: {mimeType, data}`); it becomes an MCP image block. */
interface ImagePayload {
  mimeType: string;
  data: string;
}

function ok(data: unknown): CallToolResult {
  const image = (data as { image?: ImagePayload } | null)?.image;
  if (image?.data && image.mimeType) {
    const { image: _omit, ...rest } = data as Record<string, unknown>;
    return {
      content: [
        { type: "image", data: image.data, mimeType: image.mimeType },
        { type: "text", text: JSON.stringify(rest) },
      ],
    };
  }
  return { content: [{ type: "text", text: JSON.stringify(data ?? null) }] };
}

export interface ExcelMcpOptions {
  /** Expose opt-in tools such as execute_office_js. */
  allowOptInTools?: boolean;
}

function fail(error: ToolError): CallToolResult {
  return {
    isError: true,
    content: [{ type: "text", text: JSON.stringify({ error }) }],
  };
}

/**
 * Build the Excel MCP server. Every tool call is validated against the shared
 * protocol schemas and forwarded to the bridge, which routes it to the task
 * pane holding the live workbook.
 */
export function createExcelMcpServer(forwarder: ToolForwarder, options: ExcelMcpOptions = {}): Server {
  const exposed = new Set<string>(exposedToolNames(options.allowOptInTools ?? false));
  const server = new Server(SERVER_INFO, { capabilities: { tools: {} } });

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: ALL_TOOL_DEFINITIONS.filter((tool) => exposed.has(tool.name)).map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema as Record<string, unknown>,
      annotations: tool.annotations,
    })),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request): Promise<CallToolResult> => {
    const name = request.params.name;
    if (!isExcelToolName(name) || !exposed.has(name)) {
      return fail({ code: "unknown_tool", message: `unknown tool '${name}'` });
    }

    let parsed: unknown;
    try {
      parsed = parseToolInput(name, request.params.arguments);
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
    return result.ok ? ok(result.data) : fail(result.error);
  });

  return server;
}
