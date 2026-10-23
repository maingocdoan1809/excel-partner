import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ToolResult } from "../protocol/src/index.js";
import type { SavedReferenceStatus } from "../bridge/src/saved-references.js";
import { SavedReferences } from "../bridge/src/saved-references.js";
import { ReferenceFiles } from "../bridge/src/reference-files.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput } from "../providers/src/types.js";
import { PaneClient, getFreePort } from "./helpers.js";

const folders: string[] = [];
let bridge: Bridge | undefined;
const panes: PaneClient[] = [];
afterEach(async () => {
  await Promise.all(panes.splice(0).map(pane => pane.close()));
  await bridge?.close();
  bridge = undefined;
  for (const folder of folders.splice(0)) fs.rmSync(folder, { recursive: true, force: true });
});
function fixture() {
  const folder = fs.mkdtempSync(path.join(os.tmpdir(), "saved-references-"));
  folders.push(folder);
  const file = path.join(folder, "knowledge.md");
  fs.writeFileSync(file, "Initial knowledge");
  return { folder, file, metadata: path.join(folder, "settings", "references.json") };
}

describe("saved reference files", () => {
  it("persists only metadata, reads current content and revokes exact file access on removal", () => {
    const { file, folder, metadata } = fixture();
    const store = new SavedReferences(metadata);
    const entry = store.save({ name: "Knowledge", path: file, description: "Business rules" });
    const restored = new SavedReferences(metadata);
    expect(restored.list()).toEqual([{ ...entry, available: true }]);
    expect(fs.readFileSync(metadata, "utf8")).not.toContain("Initial knowledge");
    const access = new ReferenceFiles([]);
    access.setSavedPaths(() => restored.paths());
    fs.writeFileSync(file, "Updated knowledge");
    expect(access.readSelectedFile(file).bytes.toString()).toBe("Updated knowledge");
    const sibling = path.join(folder, "private.md");
    fs.writeFileSync(sibling, "Private");
    expect(() => access.resolve(sibling)).toThrow("outside");
    expect(() => access.resolve(folder)).toThrow("outside");
    restored.remove(entry.id);
    expect(() => access.resolve(file)).toThrow("outside");
    expect(fs.readFileSync(file, "utf8")).toBe("Updated knowledge");
  });

  it("retains missing links for editing and rejects duplicates", () => {
    const { file, metadata } = fixture();
    const store = new SavedReferences(metadata);
    const entry = store.save({ name: "Knowledge", path: file, description: "Rules" });
    expect(() => store.save({ name: "Again", path: file, description: "" })).toThrow("already saved");
    fs.unlinkSync(file);
    expect(store.list()[0]?.available).toBe(false);
    expect(store.prompt()).toContain('"available":false');
    store.save({ ...entry, description: "Updated description" });
    expect(new SavedReferences(metadata).list()[0]?.description).toBe("Updated description");
  });

  it("does not follow a registered path replaced by a symlink to an unregistered file", () => {
    if (process.platform === "win32") return;
    const { file, folder } = fixture();
    const store = new SavedReferences(null);
    store.save({ name: "Knowledge", path: file, description: "" });
    const access = new ReferenceFiles([]);
    access.setSavedPaths(() => store.paths());
    const target = path.join(folder, "secret.md");
    fs.writeFileSync(target, "Secret");
    fs.unlinkSync(file);
    fs.symlinkSync(target, file);
    expect(() => access.resolve(file)).toThrow("outside");
  });

  it("validates typed paths, exposes metadata to AI and reads originals through MCP", async () => {
    const { file, metadata } = fixture();
    let prompt = "";
    class CaptureAdapter extends MockAdapter {
      override async *sendTurn(input: SendTurnInput) { prompt = input.prompt; yield* super.sendTurn(input); }
    }
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    bridge = new Bridge(config, { savedReferencesFile: metadata, references: new ReferenceFiles([]),
      usageFile: null, registry: new ProviderRegistry().register(new CaptureAdapter({ id: "claude" })) });
    const { url } = await bridge.listen();
    async function connect(id: string) {
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      panes.push(pane);
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: id });
      const hello = await pane.waitFor(m => m.type === "hello_ok");
      if (hello.type !== "hello_ok") throw new Error("No hello");
      return { pane, headers: { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" } };
    }
    const first = await connect("reference-first-123");
    const second = await connect("reference-second-123");
    const request = (route: string, method = "GET", body?: unknown, headers = first.headers) => fetch(`${url}${route}`, { method, headers, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
    expect((await fetch(`${url}/session/references`)).status).toBe(401);
    expect((await request("/session/references", "POST", { name: "Knowledge", description: "Rules", path: path.join(path.dirname(file), "missing.md") })).status).toBe(400);
    expect((await request("/session/references", "POST", { name: "Knowledge", description: "Rules", path: path.dirname(file) })).status).toBe(400);
    const input = { name: "Knowledge", description: "Common rules", path: `  "${file}"  ` };
    expect((await request("/session/references", "POST", input, second.headers)).status).toBe(200);
    expect((await request("/session/references", "DELETE", { id: ((await (await request("/session/references")).json()) as { references: SavedReferenceStatus[] }).references[0]!.id })).status).toBe(200);
    expect((await request("/session/references", "POST", input)).status).toBe(200);
    const entries = await (await request("/session/references")).json() as { references: SavedReferenceStatus[] };
    const id = entries.references[0]!.id;
    expect(entries.references[0]).toMatchObject({ name: "Knowledge", available: true });
    expect((await request("/session/references", "POST", input)).status).toBe(400);
    first.pane.send({ type: "set_provider", provider: "claude" });
    await first.pane.waitFor(m => m.type === "provider_status");
    first.pane.send({ type: "prompt", requestId: "references-turn", prompt: "Consult the rules", selection: null });
    await first.pane.waitFor(m => m.type === "turn_finished");
    expect(prompt).toContain("<saved_reference_files>");
    expect(prompt).toContain("Common rules");
    expect(prompt).not.toContain("Initial knowledge");
    fs.writeFileSync(file, "Latest knowledge");
    const result = await (await request("/internal/tool", "POST", { tool: "read_attached_file", args: { path: file } })).json() as ToolResult;
    expect(result.ok).toBe(true);
    expect(JSON.stringify(result)).toContain("Latest knowledge");
    expect((await request("/session/references", "POST", { id, name: "Renamed", description: "Edited" })).status).toBe(200);
    const replacement = path.join(path.dirname(file), "new-common.md");
    fs.writeFileSync(replacement, "Replacement rules");
    expect((await request("/session/references", "POST", { id, path: replacement, name: "Renamed", description: "Relinked" })).status).toBe(200);
    let accessRequests = 0;
    first.pane.onMessage = message => {
      if (message.type === "tool_request" && message.tool === "__authorize_reference_path") {
        accessRequests++;
        first.pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { approved: false } } });
      }
    };
    const oldAccess = await (await request("/internal/tool", "POST", { tool: "read_attached_file", args: { path: file } })).json() as ToolResult;
    expect(oldAccess.ok).toBe(false);
    const newAccess = await (await request("/internal/tool", "POST", { tool: "read_attached_file", args: { path: replacement } })).json() as ToolResult;
    expect(newAccess.ok).toBe(true);
    expect(JSON.stringify(newAccess)).toContain("Replacement rules");
    expect((await request("/session/references", "DELETE", { id })).status).toBe(200);
    const denied = await (await request("/internal/tool", "POST", { tool: "read_attached_file", args: { path: file } })).json() as ToolResult;
    expect(denied.ok).toBe(false);
    expect(accessRequests).toBe(2);
    expect(fs.readFileSync(file, "utf8")).toBe("Latest knowledge");
  });
});
