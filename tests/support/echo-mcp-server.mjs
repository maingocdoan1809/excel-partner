// A tiny stdio MCP server for tests: one tool that answers, one that fails and one that never replies.
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";

const server = new Server({ name: "echo-fixture", version: "1.0.0" }, { capabilities: { tools: {} } });
server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    { name: "echo", description: "Echo the text back", inputSchema: { type: "object", properties: { text: { type: "string" } }, required: ["text"] } },
    { name: "fail", description: "Always fails", inputSchema: { type: "object", properties: {} } },
    { name: "hang", description: "Never answers", inputSchema: { type: "object", properties: {} } },
  ],
}));
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  if (name === "echo") return { content: [{ type: "text", text: `echo:${args?.text}:${process.env.FIXTURE_VALUE ?? ""}` }] };
  if (name === "fail") return { isError: true, content: [{ type: "text", text: "it failed" }] };
  return new Promise(() => {});
});
await server.connect(new StdioServerTransport());
