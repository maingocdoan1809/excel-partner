// @vitest-environment jsdom
import { afterEach, expect, it, vi } from "vitest";
import type { BridgeClient } from "../frontend/taskpane/bridge-client.js";
import { setupMcp } from "../frontend/taskpane/ui/mcp.js";
import { discoverMcpCatalog } from "../shared/src/mcp-catalog.js";
import { McpServers } from "../shared/src/mcp-servers.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const server = { id: "demo", name: "Fixture MCP", transport: "stdio", target: "node", args: ["demo.mjs"] };
afterEach(() => { sessionStorage.clear(); document.body.innerHTML = ""; });
function ui() {
  document.body.dataset.host = "excel";
  document.body.innerHTML = '<div><section id="settings-tools-section"><div id="tools-menu"></div></section></div>';
  const enabled = new Set<string>();
  const config = { ...server };
  const request = vi.fn(async (method = "GET", body?: any) => {
    if (method === "GET") return { servers: [config], enabled: [...enabled] };
    if (body.action === "select") { if (body.enabled) enabled.add(body.id); else enabled.delete(body.id); }
    return { ok: true };
  });
  return { settings: setupMcp(() => ({ mcpRequest: request }) as unknown as BridgeClient), enabled, config, request };
}
it("restores only explicitly selected servers after reconnecting in the same tab", async () => {
  const h = ui(); await h.settings.restore(); expect(h.enabled.size).toBe(0);
  await h.settings.refresh();
  const checkbox = document.querySelector<HTMLInputElement>('.mcp-picker input')!;
  checkbox.checked = true; checkbox.dispatchEvent(new Event('change'));
  await h.settings.ready();
  expect(h.enabled.has('demo')).toBe(true);
  h.enabled.clear(); await h.settings.restore();
  expect(h.enabled.has('demo')).toBe(true);
  h.enabled.clear(); h.config.target = 'changed-command'; await h.settings.restore();
  expect(h.enabled.size).toBe(0);
});
it("the send barrier waits until a selection is saved", async () => {
  const h = ui(); await h.settings.refresh();
  let release!: () => void;
  h.request.mockImplementationOnce(async () => { await new Promise<void>(resolve => { release = resolve; }); return { ok: true }; });
  const checkbox = document.querySelector<HTMLInputElement>('.mcp-picker input')!;
  checkbox.checked = true; checkbox.dispatchEvent(new Event('change'));
  let done = false; const ready = h.settings.ready().then(() => { done = true; });
  await Promise.resolve(); expect(done).toBe(false);
  await vi.waitFor(() => expect(release).toBeTypeOf("function"));
  release(); await ready; expect(done).toBe(true);
});
it("registers actual tool schemas only for enabled servers", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-catalog-"));
  const manager = new McpServers(path.join(dir, "servers.json"));
  try {
    await manager.save({ ...server, target: process.execPath, args: [path.resolve("tests/support/mcp-fixture.mjs")] });
    expect((await discoverMcpCatalog(manager, "pane", new Set())).size).toBe(0);
    const catalog = await discoverMcpCatalog(manager, "pane", new Set(["demo"]));
    const summary = [...catalog.values()].find(tool => tool.toolName === "echo")!;
    expect(summary.definition.name).toMatch(/__echo$/);
    expect(summary.definition.inputSchema.properties).toHaveProperty("text");
    expect(summary.definition.description).toContain("Fixture MCP");
    expect([...catalog.keys()]).not.toContain("mcp_call_tool");
    expect((await discoverMcpCatalog(manager, "pane", new Set())).size).toBe(0);
  } finally { await manager.close(); fs.rmSync(dir, { recursive: true, force: true }); }
});

it("a failed server selection does not block subsequent chat messages", async () => {
  const h = ui(); await h.settings.refresh();
  h.request.mockRejectedValueOnce(new Error("Server unavailable"));
  const checkbox = document.querySelector<HTMLInputElement>('.mcp-picker input')!;
  checkbox.checked = true; checkbox.dispatchEvent(new Event('change'));
  await vi.waitFor(() => expect(checkbox.disabled).toBe(false));
  expect(checkbox.checked).toBe(false);
  expect(document.querySelector('.mcp-status')!.textContent).toContain('Server unavailable');
  await expect(h.settings.ready()).resolves.toBeUndefined();
});

it("an unavailable remembered server does not block Office chat after reconnect", async () => {
  const h = ui(); await h.settings.refresh();
  const checkbox = document.querySelector<HTMLInputElement>('.mcp-picker input')!;
  checkbox.checked = true; checkbox.dispatchEvent(new Event('change')); await h.settings.ready();
  h.enabled.clear();
  h.request.mockResolvedValueOnce({ servers: [h.config], enabled: [] }).mockRejectedValueOnce(new Error('Remembered server offline'));
  await h.settings.restore();
  expect(document.querySelector('.mcp-status')!.textContent).toContain('Remembered server offline');
  await expect(h.settings.ready()).resolves.toBeUndefined();
});
