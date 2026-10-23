import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { SkillStore } from "../bridge/src/skills.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent, ToolResult } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

class BoundaryAdapter extends MockAdapter {
  contexts: StartSessionContext[] = [];
  inputs: SendTurnInput[] = [];
  results: ToolResult[] = [];
  useTool = true;
  release!: () => void;
  private gate = new Promise<void>(resolve => { this.release = resolve; });
  constructor() { super({ id: "claude" }); }
  override async startSession(context: StartSessionContext) {
    this.contexts.push(context);
    return super.startSession(context);
  }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    this.inputs.push(input);
    yield { type: "session_info", requestId: input.requestId, provider: "claude", providerSessionId: "native-session" };
    yield { type: "turn_started", requestId: input.requestId };
    await this.gate;
    if (this.useTool && !input.signal.aborted) {
      const mcp = this.contexts.at(-1)!.mcp;
      const response = await fetch(`${mcp.bridgeUrl}/internal/tool`, {
        method: "POST", headers: { authorization: `Bearer ${mcp.token}`, "content-type": "application/json" },
        body: JSON.stringify({ sessionId: mcp.sessionId, tool: "get_sheets", args: {} }),
      });
      this.results.push(await response.json() as ToolResult);
    }
    yield { type: "turn_finished", requestId: input.requestId, reason: input.signal.aborted ? "cancelled" : "completed" };
  }
}

let bridge: Bridge | null = null;
let pane: PaneClient | null = null;
let folder = "";
afterEach(async () => {
  await pane?.close(); pane = null;
  await bridge?.close(); bridge = null;
  if (folder) fs.rmSync(folder, { recursive: true, force: true });
});
async function setup() {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "settings-followups-"));
  const upload = path.join(folder, "uploaded");
  const external = path.join(folder, "external");
  fs.mkdirSync(path.join(external, "readonly"), { recursive: true });
  fs.writeFileSync(path.join(external, "readonly", "SKILL.md"), "---\nname: readonly\n---\nPackaged rules");
  const port = await getFreePort();
  const adapter = new BoundaryAdapter();
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
    usageFile: null, savedReferencesFile: null, workspaceConfigPath: null,
    skillUploadDir: upload, skills: new SkillStore([upload, external], path.join(folder, "templates")),
    systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(adapter),
  });
  const { url } = await bridge.listen();
  pane = await PaneClient.connect(url.replace("http", "ws") + "/ws");
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "test-workbook" });
  const hello = await pane.waitFor(message => message.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("no handshake");
  const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
  pane.onMessage = message => {
    if (message.type === "tool_request") pane!.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { sheets: [] } } });
  };
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor(message => message.type === "provider_status");
  return { url, headers, adapter, pane };
}

describe("editable settings and pending user prompts", () => {
  it("persists skill toggles and keeps disabled skills manageable while hiding them from AI", async () => {
    const { url, headers } = await setup();
    const patch = (body: unknown) => fetch(url + "/session/skills", { method: "PATCH", headers, body: JSON.stringify(body) });
    expect((await fetch(url + "/session/skills", { method: "PATCH", body: JSON.stringify({ name: "readonly", enabled: false }) })).status).toBe(401);
    expect((await patch({ name: "readonly", enabled: "false" })).status).toBe(400);
    expect((await patch({ name: "readonly", enabled: false })).status).toBe(200);
    expect(await (await fetch(url + "/skills")).json()).toEqual({ skills: [] });
    expect(await (await fetch(url + "/session/skills", { headers })).json()).toMatchObject({ skills: [{ name: "readonly", enabled: false, removable: false }] });
    const restored = new SkillStore([path.join(folder, "external")], path.join(folder, "templates"));
    expect(restored.isEnabled("readonly")).toBe(false);
    expect((await patch({ name: "readonly", enabled: true })).status).toBe(200);
    expect(restored.isEnabled("readonly")).toBe(true);
    expect(await (await fetch(url + "/skills")).json()).toMatchObject({ skills: [{ name: "readonly" }] });
  });

  it("updates provider tool exposure on the next message and resumes the native conversation", async () => {
    const { pane, adapter } = await setup();
    adapter.useTool = false;
    adapter.release();
    pane.send({ type: "prompt", requestId: "tools-off", prompt: "Hello", selection: null });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "tools-off");
    expect(adapter.contexts[0]!.mcp.env.EXCEL_WEB_SEARCH).toBeUndefined();
    pane.send({ type: "prompt", requestId: "tools-on", prompt: "Search", selection: null, resumeSessionId: "native-session", capabilities: { webSearch: true, playwright: true } });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "tools-on");
    expect(adapter.contexts[1]).toMatchObject({ resumeSessionId: "native-session", mcp: { env: { EXCEL_WEB_SEARCH: "1", EXCEL_PLAYWRIGHT: "1" } } });
    pane.send({ type: "prompt", requestId: "tools-off-again", prompt: "Continue", selection: null, resumeSessionId: "native-session" });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "tools-off-again");
    expect(adapter.contexts).toHaveLength(3);
    expect(adapter.contexts[2]!.mcp.env.EXCEL_PLAYWRIGHT).toBeUndefined();
  });

  it("authenticates skill editing, rejects stale and non-owned files, and refreshes skill metadata", async () => {
    const { url, headers } = await setup();
    expect((await fetch(url + "/session/skills?name=report")).status).toBe(401);
    const manifest = "---\nname: report\ndescription: Original\n---\nOriginal instructions";
    expect((await fetch(url + "/session/skills", { method: "POST", headers: { ...headers, "x-file-name": "report.md" }, body: manifest })).status).toBe(201);
    expect(await (await fetch(url + "/session/skills?name=report", { headers })).json()).toEqual({ files: ["SKILL.md"] });
    const { document } = await (await fetch(url + "/session/skills?name=report&file=SKILL.md", { headers })).json() as { document: { content: string; revision: string; file: string } };
    const body = JSON.stringify({ name: "report", ...document, content: manifest.replaceAll("Original", "Updated") });
    expect((await fetch(url + "/session/skills", { method: "PUT", headers, body })).status).toBe(200);
    expect((await fetch(url + "/session/skills", { method: "PUT", headers, body })).status).toBe(409);
    expect(await (await fetch(url + "/skills")).json()).toMatchObject({ skills: expect.arrayContaining([{ name: "report", description: "Updated" }]) });
    expect((await fetch(url + "/session/skills?name=readonly&file=SKILL.md", { headers })).status).toBe(403);
    expect((await fetch(url + "/session/skills", { method: "PUT", headers, body: JSON.stringify({ name: "readonly", ...document }) })).status).toBe(403);
  });

  it("delivers additional user instructions once at the next tool boundary in order", async () => {
    const { adapter, pane } = await setup();
    pane.send({ type: "prompt", requestId: "first", prompt: "Make a report", selection: null });
    await pane.waitFor(message => message.type === "turn_started");
    for (const [id, prompt] of [["one", "Use blue headers"], ["two", "And keep formulas"]]) {
      pane.send({ type: "followup", requestId: "first", id: id!, prompt: prompt! });
      await pane.waitFor(message => message.type === "followup_status" && message.id === id && message.state === "pending");
    }
    adapter.release();
    await pane.waitFor(message => message.type === "turn_finished");
    expect(adapter.results[0]).toMatchObject({ ok: true, additional_user_messages: [{ id: "one", text: "Use blue headers" }, { id: "two", text: "And keep formulas" }] });
    expect(pane.received.filter(message => message.type === "followup_status" && message.state === "delivered")).toHaveLength(2);
    pane.send({ type: "prompt", requestId: "second", prompt: "Continue", selection: null, resumeSessionId: "native-session" });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "second");
    expect(adapter.results[1]?.additional_user_messages).toBeUndefined();
  });

  it("returns unsent followups for the next turn and does not lose late arrivals", async () => {
    const { adapter, pane } = await setup();
    adapter.useTool = false;
    pane.send({ type: "prompt", requestId: "first", prompt: "Answer", selection: null });
    await pane.waitFor(message => message.type === "turn_started");
    pane.send({ type: "followup", requestId: "first", id: "queued", prompt: "Also explain" });
    await pane.waitFor(message => message.type === "followup_status" && message.state === "pending");
    adapter.release();
    await pane.waitFor(message => message.type === "turn_finished");
    expect(pane.received).toContainEqual({ type: "followup_status", requestId: "first", id: "queued", state: "next_turn" });
    pane.send({ type: "followup", requestId: "first", id: "late", prompt: "Late message" });
    await pane.waitFor(message => message.type === "followup_status" && message.id === "late" && message.state === "next_turn");
    expect(adapter.results).toEqual([]);
  });

  it("keeps uploaded attachments with a pending message and rejects missing attachments", async () => {
    const { adapter, pane, url, headers } = await setup();
    const uploaded = await (await fetch(url + "/session/attachment", { method: "POST", headers: { ...headers, "x-file-name": "notes.md" }, body: "# Notes" })).json() as { id: string; path: string };
    pane.send({ type: "prompt", requestId: "first", prompt: "Work", selection: null });
    await pane.waitFor(message => message.type === "turn_started");
    pane.send({ type: "followup", requestId: "first", id: "with-file", prompt: "Use @notes.md", attachmentIds: [uploaded.id] });
    await pane.waitFor(message => message.type === "followup_status" && message.id === "with-file" && message.state === "pending");
    pane.send({ type: "followup", requestId: "first", id: "missing", prompt: "Missing file", attachmentIds: ["00000000-0000-4000-8000-000000000000"] });
    await pane.waitFor(message => message.type === "followup_status" && message.id === "missing" && message.state === "rejected");
    adapter.release();
    await pane.waitFor(message => message.type === "turn_finished");
    const followups = adapter.results[0]?.additional_user_messages;
    expect(followups).toHaveLength(1);
    expect(followups?.[0]?.text).toContain(uploaded.path);
    expect(followups?.[0]?.text).toContain("Use @notes.md");
  });

  it("applies a system prompt edit on the next request while retaining the native session", async () => {
    const { adapter, pane, url, headers } = await setup();
    adapter.useTool = false;
    pane.send({ type: "prompt", requestId: "first", prompt: "Hello", selection: null });
    await pane.waitFor(message => message.type === "turn_started");
    expect((await fetch(url + "/session/system-prompt")).status).toBe(401);
    const initial = await (await fetch(url + "/session/system-prompt", { headers })).json() as { revision: string; content: string };
    expect(initial.content).toBe("");
    const content = "Use the Excel tools and answer in Vietnamese.";
    const response = await fetch(url + "/session/system-prompt", { method: "PUT", headers, body: JSON.stringify({ revision: initial.revision, content }) });
    expect(response.status).toBe(200);
    expect(adapter.inputs[0]?.systemPrompt).toBe(initial.content);
    adapter.release();
    await pane.waitFor(message => message.type === "turn_finished");
    pane.send({ type: "prompt", requestId: "second", prompt: "Next", selection: null, resumeSessionId: "native-session" });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "second");
    expect(adapter.contexts).toHaveLength(2);
    expect(adapter.contexts[1]).toMatchObject({ systemPrompt: content, resumeSessionId: "native-session" });
    expect(adapter.inputs[1]?.systemPrompt).toBe(content);
  });

  it("does not deliver queued messages after Stop or cancel a new request for a stale cancel ID", async () => {
    const { adapter, pane } = await setup();
    pane.send({ type: "prompt", requestId: "first", prompt: "Work", selection: null });
    await pane.waitFor(message => message.type === "turn_started");
    pane.send({ type: "cancel", requestId: "older-request" });
    pane.send({ type: "followup", requestId: "first", id: "queued", prompt: "Pending change" });
    await pane.waitFor(message => message.type === "followup_status" && message.state === "pending");
    expect(adapter.inputs[0]?.signal.aborted).toBe(false);
    pane.send({ type: "cancel", requestId: "first" });
    // Wait for the server's cancellation before releasing the simulated provider.
    pane.send({ type: "followup", requestId: "first", id: "after-stop", prompt: "Keep this for later" });
    await pane.waitFor(message => message.type === "followup_status" && message.id === "after-stop" && message.state === "next_turn");
    adapter.release();
    const finished = await pane.waitFor(message => message.type === "turn_finished");
    expect(finished).toMatchObject({ reason: "cancelled" });
    expect(adapter.results).toEqual([]);
    expect(pane.received.some(message => message.type === "followup_status" && message.state === "delivered")).toBe(false);
    expect(pane.received).toContainEqual({ type: "followup_status", requestId: "first", id: "queued", state: "next_turn" });
  });
});
