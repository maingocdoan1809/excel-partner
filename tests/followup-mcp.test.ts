import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createExcelMcpServer } from "../excel_mcp/src/server.js";

describe("pending prompts through MCP", () => {
  it.each([true, false])("includes direct user followups even when the tool succeeded=%s", async succeeded => {
    const server = createExcelMcpServer({ call: async () => ({
      ...(succeeded ? { ok: true as const, data: { sheets: [] } } : { ok: false as const, error: { code: "office_error" as const, message: "Cannot read sheet" } }),
      additional_user_messages: [{ id: "followup", text: "Use the Summary tab instead" }],
    }) });
    const client = new Client({ name: "test", version: "1" });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    try {
      await server.connect(serverTransport);
      await client.connect(clientTransport);
      const result = await client.callTool({ name: "get_sheets", arguments: {} });
      const blocks = result.content as Array<{ type: string; text?: string }>;
      expect(blocks).toHaveLength(2);
      expect(blocks[1]?.text).toContain("Use the Summary tab instead");
      expect(blocks[1]?.text).toContain("Additional messages sent directly by the user");
      expect(result.isError ?? false).toBe(!succeeded);
    } finally { await client.close(); await server.close(); }
  });
});
