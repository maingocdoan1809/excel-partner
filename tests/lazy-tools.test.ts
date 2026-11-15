import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createExcelMcpServer } from "../excel_mcp/src/server.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { HOST_IDS, parseToolInputForHost, toolGroupsFromEnv, type HostId, type ToolGroups } from "../protocol/src/index.js";
import { describableTools, officeToolCatalog, toolCatalogChars } from "../shared/src/catalog.js";
import {
  GET_TOOL_DEFINITION, HOT_TOOLS, MAX_DEFINITIONS, MAX_SUMMARY_CHARS, TOOL_SUMMARIES, firstSentence, lookupTools, normalizeToolName, suggestToolNames, summaryOf,
} from "../shared/src/lazy-tools.js";
import { PaneClient, getFreePort } from "./helpers.js";

const EVERYTHING: ToolGroups = { webSearch: true, playwright: true, connectedFiles: true };
const LAZY: ToolGroups = { ...EVERYTHING, lazySchemas: true };
const full = (host: HostId) => describableTools(host, true, EVERYTHING).filter(tool => tool.name !== GET_TOOL_DEFINITION);

describe("one-line summaries", () => {
  it.each(HOST_IDS)("%s: every tool has a clean summary within the limit", host => {
    for (const tool of full(host)) {
      const summary = summaryOf(tool);
      expect(summary.length, tool.name).toBeLessThanOrEqual(MAX_SUMMARY_CHARS);
      expect(summary.endsWith("…"), `${tool.name} has no clean first sentence: write a TOOL_SUMMARIES entry`).toBe(false);
      expect(TOOL_SUMMARIES[tool.name] !== undefined || firstSentence(tool.description) !== null, tool.name).toBe(true);
    }
  });

  it("keeps the access tag and does not stop at an abbreviation or a version number", () => {
    expect(firstSentence("READ. Values of a range (e.g. A1:B2) in ExcelApi 1.19+. More text.")).toBe("READ. Values of a range (e.g. A1:B2) in ExcelApi 1.19+.");
    expect(firstSentence("WRITE (operation list only reads). Do a thing. Then more.")).toBe("WRITE (operation list only reads). Do a thing.");
    expect(firstSentence("x".repeat(400))).toBeNull();
  });

  it("has no summary for a tool that does not exist", () => {
    const names = new Set(HOST_IDS.flatMap(host => full(host).map(tool => tool.name)));
    for (const name of Object.keys(TOOL_SUMMARIES)) expect(names.has(name), `stale summary: ${name}`).toBe(true);
  });

  it("lists only tools that exist among the hot tools", () => {
    for (const host of HOST_IDS) for (const name of HOT_TOOLS[host]) expect(full(host).some(tool => tool.name === name), `${host}: ${name}`).toBe(true);
  });
});

describe("lazy catalog", () => {
  it.each(HOST_IDS)("%s: stubs everything but the hot tools, keeps annotations, and offers get_tool_definition", host => {
    const lazy = officeToolCatalog(host, true, LAZY);
    const fullCatalog = full(host);
    expect(lazy.map(tool => tool.name).sort()).toEqual([...fullCatalog.map(tool => tool.name), GET_TOOL_DEFINITION].sort());
    for (const tool of lazy) {
      const real = describableTools(host, true, EVERYTHING).find(candidate => candidate.name === tool.name)!;
      if (tool.name === GET_TOOL_DEFINITION || HOT_TOOLS[host].includes(tool.name)) {
        expect(tool.inputSchema, tool.name).toEqual(real.inputSchema);
        expect(tool.description, tool.name).toBe(real.description);
      } else {
        expect(tool.inputSchema, tool.name).toEqual({ type: "object", additionalProperties: true });
        expect(tool.description, tool.name).toContain(`Call ${GET_TOOL_DEFINITION} before first use.`);
        expect(tool.annotations, tool.name).toEqual(real.annotations);
      }
    }
  });

  it("is unchanged, and has no get_tool_definition, when schemas are full", () => {
    const names = officeToolCatalog("excel", true, EVERYTHING).map(tool => tool.name);
    expect(names).not.toContain(GET_TOOL_DEFINITION);
    expect(officeToolCatalog("excel", true, EVERYTHING).every(tool => (tool.inputSchema as { additionalProperties?: unknown }).additionalProperties !== true || tool.name === "playwright_browser")).toBe(true);
  });

  it("is much smaller than the full catalog", () => {
    for (const [host, ceiling] of [["excel", 0.35], ["word", 0.55], ["powerpoint", 0.45]] as const) {
      expect(toolCatalogChars(host, true, LAZY) / toolCatalogChars(host, true, EVERYTHING), host).toBeLessThan(ceiling);
    }
  });

  it("reads the setting from the MCP process environment", () => {
    expect(toolGroupsFromEnv({ EXCEL_TOOL_SCHEMAS: "lazy", EXCEL_PLAYWRIGHT: "1" })).toMatchObject({ lazySchemas: true, playwright: true, connectedFiles: false });
    expect(toolGroupsFromEnv({}).lazySchemas).toBe(false);
  });
});

describe("finding a tool by a name that is nearly right", () => {
  const catalog = full("excel");
  const names = catalog.map(tool => tool.name);

  it("normalises case, separators and the MCP prefix", () => {
    for (const input of ["read_range", "Read Range", "read-range", "readRange", "mcp__excel__read_range", "READ_RANGE"]) expect(normalizeToolName(input), input).toBe("read_range");
  });

  it("suggests the closest tools for a typo, a shortened name or a missing separator", () => {
    expect(suggestToolNames("read_rang", names)[0]).toBe("read_range");
    expect(suggestToolNames("readrange", names)[0]).toBe("read_range");
    expect(suggestToolNames("conditional", names)).toContain("conditional_format");
    expect(suggestToolNames("sort", names)).toContain("sort_range");
    expect(suggestToolNames("zzzzqqq", names)).toEqual([]);
    expect(suggestToolNames("", names)).toEqual([]);
  });

  it("returns definitions for names it can resolve", () => {
    const lookup = lookupTools(["Read Range", "set_cells"], catalog);
    expect(lookup.definitions.map(tool => tool.name)).toEqual(["read_range", "set_cells"]);
    expect(lookup.unmatched).toEqual([]);
    expect(lookup.tools).toBeUndefined();
    expect(lookup.definitions[0]!.inputSchema).toMatchObject({ type: "object" });
  });

  it("reports an unknown name with suggestions and keeps the ones it found", () => {
    const lookup = lookupTools(["read_range", "sort_rnge"], catalog);
    expect(lookup.definitions.map(tool => tool.name)).toEqual(["read_range"]);
    expect(lookup.unmatched).toEqual([{ input: "sort_rnge", suggestions: expect.arrayContaining(["sort_range"]) }]);
    expect(lookup.tools).toBeUndefined();
    expect(lookup.note).toContain("sort_rnge");
  });

  it("lists every tool when nothing matches, and when nothing is asked", () => {
    for (const requested of [["no_such_tool", "???"], []]) {
      const lookup = lookupTools(requested, catalog);
      expect(lookup.definitions).toEqual([]);
      expect(lookup.tools?.map(tool => tool.name)).toEqual(names);
      expect(lookup.tools?.every(tool => tool.summary.length > 0)).toBe(true);
      expect(lookup.note).toContain(GET_TOOL_DEFINITION);
    }
  });

  it("does not return the same tool twice, and limits one answer", () => {
    expect(lookupTools(["read_range", "read_range", "readRange"], catalog).definitions).toHaveLength(1);
    const lookup = lookupTools(names, catalog);
    expect(lookup.definitions).toHaveLength(MAX_DEFINITIONS);
    expect(lookup.note).toContain("left out");
  });

  it("accepts a lone string, a wrong type, and extra parameters without failing", () => {
    expect(parseToolInputForHost("excel", GET_TOOL_DEFINITION, { names: "read_range" })).toMatchObject({ names: ["read_range"] });
    expect(parseToolInputForHost("excel", GET_TOOL_DEFINITION, { names: 7, other: true })).toMatchObject({ other: true });
    expect(parseToolInputForHost("excel", GET_TOOL_DEFINITION, {})).toEqual({});
  });
});

async function connect(capabilities: ToolGroups) {
  const forward = vi.fn(async () => ({ ok: true as const, data: { forwarded: true } }));
  const server = createExcelMcpServer({ call: forward }, { allowOptInTools: true, capabilities });
  const client = new Client({ name: "test", version: "1" });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  return { client, forward, close: async () => { await client.close(); await server.close(); } };
}
const textOf = (result: unknown) => ((result as { content: Array<{ text: string }> }).content[0]!.text);

describe("MCP server with lazy schemas", () => {
  it("lists stubs, and a call with bad arguments gets the schema back", async () => {
    const { client, forward, close } = await connect(LAZY);
    try {
      const tools = (await client.listTools()).tools;
      expect(tools.find(tool => tool.name === "read_range")!.inputSchema).toMatchObject({ properties: expect.any(Object) });
      expect(tools.find(tool => tool.name === "set_cells")!.inputSchema).toEqual({ type: "object", additionalProperties: true });
      const bad = await client.callTool({ name: "set_cells", arguments: {} });
      expect(bad.isError).toBe(true);
      expect(textOf(bad)).toContain("Input schema of set_cells");
      expect(textOf(bad)).toContain("\"cells\"");
      expect(forward).not.toHaveBeenCalled();
    } finally { await close(); }
  });

  it("does not repeat a schema the model already has, in full mode or for a hot tool", async () => {
    for (const [capabilities, tool] of [[LAZY, "read_range"], [EVERYTHING, "set_cells"]] as const) {
      const { client, close } = await connect(capabilities);
      try { expect(textOf(await client.callTool({ name: tool, arguments: { bogus: 1 } }))).not.toContain("Input schema of"); }
      finally { await close(); }
    }
  });

  it("suggests the right name when a tool does not exist", async () => {
    const { client, forward, close } = await connect(LAZY);
    try {
      const result = await client.callTool({ name: "read_rnge", arguments: {} });
      expect(result.isError).toBe(true);
      expect(textOf(result)).toContain("Did you mean: read_range");
      expect(textOf(await client.callTool({ name: "zzzzqqq", arguments: {} }))).toContain(`Call ${GET_TOOL_DEFINITION} with no names`);
      expect(forward).not.toHaveBeenCalled();
    } finally { await close(); }
  });

  it("forwards get_tool_definition to the bridge, and offers it only in lazy mode", async () => {
    const lazy = await connect(LAZY);
    try {
      await lazy.client.callTool({ name: GET_TOOL_DEFINITION, arguments: { names: ["set_cells"] } });
      expect(lazy.forward).toHaveBeenCalledWith(GET_TOOL_DEFINITION, expect.objectContaining({ names: ["set_cells"] }));
    } finally { await lazy.close(); }
    const regular = await connect(EVERYTHING);
    try { expect((await regular.client.listTools()).tools.map(tool => tool.name)).not.toContain(GET_TOOL_DEFINITION); }
    finally { await regular.close(); }
  });
});

describe("get_tool_definition through the bridge", () => {
  it("is answered by the bridge itself, without the task pane, and only when the bridge is lazy", async () => {
    const folder = fs.mkdtempSync(path.join(os.tmpdir(), "lazy-tools-"));
    const file = path.join(folder, "usage.jsonl");
    const port = await getFreePort();
    const config = { ...loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), toolSchemas: "lazy" as const };
    const bridge = new Bridge(config, { usageFile: file, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", toolCalls: [
      { tool: GET_TOOL_DEFINITION, args: { names: ["set_cells", "sort_rnge"] } }] })) });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    try {
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "lazy-tools-workbook" });
      await pane.waitFor(m => m.type === "hello_ok");
      pane.send({ type: "set_provider", provider: "claude" });
      await pane.waitFor(m => m.type === "provider_status");
      pane.send({ type: "prompt", requestId: "lazy-1", chatSessionId: "chat-lazy", prompt: "describe", selection: null });
      await pane.waitFor(m => m.type === "turn_finished" && m.requestId === "lazy-1");
      expect(pane.received.filter(m => m.type === "tool_request")).toEqual([]);
      expect(pane.received.find(m => m.type === "tool_finished")).toMatchObject({ tool: GET_TOOL_DEFINITION, ok: true });
      const entry = fs.readFileSync(file, "utf8").split("\n").filter(Boolean).map(line => JSON.parse(line)).find(line => line.requestId === "lazy-1");
      expect(entry.tools[GET_TOOL_DEFINITION].calls).toBe(1);
      expect(entry.catalogChars).toBe(toolCatalogChars("excel", true, { ...EVERYTHING, playwright: false, webSearch: false, connectedFiles: false, lazySchemas: true }));
    } finally {
      await pane.close();
      await bridge.close();
      fs.rmSync(folder, { recursive: true, force: true });
    }
  });
});
