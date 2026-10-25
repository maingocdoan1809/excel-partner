import { describe, expect, it } from "vitest";
import { clientToolResult } from "../gateway/src/tool-result.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { CallToolRequestSchema, CallToolResultSchema } from "@modelcontextprotocol/sdk/types.js";
import { gatewayProbeImage } from "./support/gateway-probe-image.js";

const png = gatewayProbeImage().toString("base64");

describe("API client result to MCP", () => {
  it("keeps the application envelope intact and exposes it as a structured object", () => {
    const text = JSON.stringify({ success: true, result: { sheets: ["Sheet1", "Sheet4"] } });
    expect(clientToolResult(text, false)).toEqual({
      content: [{ type: "text", text }], isError: false,
      structuredContent: { success: true, result: { sheets: ["Sheet1", "Sheet4"] } },
    });
  });
  it.each(["ItemAlreadyExists", "InvalidOperation", "PropertyNotLoaded"])("preserves the original Excel error (%s)", code => {
    const value = { success: false, code, error: "original client error" };
    const content = [{ type: "text", text: JSON.stringify(value) }];
    const result = clientToolResult(content, true);
    expect(result.content).toEqual(content);
    expect(result.isError).toBe(true);
    expect(result.structuredContent).toEqual(value);
  });
  it("does not rewrite plain text, arrays, multiple blocks or JSON with extra text", () => {
    for (const text of ["plain result", "[1,2]", '{"result":1}\n[id:example]']) {
      expect(clientToolResult(text, false)).toEqual({ content: [{ type: "text", text }], isError: false });
    }
    const content = [{ type: "text", text: '{"a":1}' }, { type: "text", text: '{"b":2}' }];
    expect(clientToolResult(content, false)).toEqual({ content, isError: false });
  });
  it("converts Anthropic image sources without changing API history or dropping accompanying data", () => {
    const original = [
      { type: "text", text: '{"success":true,"result":{"chart":"QA"}}' },
      { type: "image", source: { type: "base64", media_type: "image/png", data: png } },
    ];
    const snapshot = JSON.stringify(original);
    const result = clientToolResult(original, false);
    expect(result.content).toEqual([original[0], { type: "image", mimeType: "image/png", data: png }]);
    expect(result.structuredContent).toEqual({ success: true, result: { chart: "QA" } });
    expect(CallToolResultSchema.safeParse(result).success).toBe(true);
    expect(JSON.stringify(original)).toBe(snapshot);
  });
  it("keeps existing MCP images and the original client error marker", () => {
    const image = { type: "image", mimeType: "image/png", data: png };
    expect(clientToolResult([image], true)).toEqual({ content: [image], isError: true });
  });
  it.each([
    [{ type: "image", source: { type: "url", url: "https://example.invalid/image" } }],
    [{ type: "image", source: { type: "base64", media_type: "image/png" } }],
    [{ type: "image" }],
    [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: png } }],
    [null],
  ])("turns unsupported/malformed media into a valid, actionable MCP error (%j)", block => {
    const result = clientToolResult([block], false);
    expect(result.isError).toBe(true);
    expect(result._meta?.["agent-gateway/conversion-error"]).toBe(true);
    expect(JSON.stringify(result.content)).toContain("may already have executed");
    expect(CallToolResultSchema.safeParse(result).success).toBe(true);
  });
  it.each([true, false])("passes the real MCP client result validator for an image, valid=%s", async valid => {
    const server = new Server({ name: "gateway-image-regression", version: "1" }, { capabilities: { tools: {} } });
    server.setRequestHandler(CallToolRequestSchema, async () => clientToolResult([
      { type: "image", source: valid ? { type: "base64", media_type: "image/png", data: png } : { type: "url", url: "https://example.invalid/image" } },
    ], false));
    const client = new Client({ name: "image-validator", version: "1" });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    try {
      await server.connect(serverTransport);
      await client.connect(clientTransport);
      const result = CallToolResultSchema.parse(await client.callTool({ name: "image_probe", arguments: {} }));
      expect(result.isError).toBe(!valid);
      expect(result.content[0]).toMatchObject(valid ? { type: "image", mimeType: "image/png", data: png } : { type: "text" });
    } finally { await client.close(); await server.close(); }
  });
});
