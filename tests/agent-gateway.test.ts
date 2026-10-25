import { afterEach, describe, expect, it, vi } from "vitest";
import { GatewayServer } from "../gateway/src/server.js";
import { ConfigSchema, routeModel } from "../gateway/src/config.js";
import { parseRequest, type ApiMessage, type ChatMessage, type ContentBlock } from "../gateway/src/protocol.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent } from "../protocol/src/index.js";
import { gatewayProbeImage } from "./support/gateway-probe-image.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const token = "agent-gateway-test-token-only";
const tools = [
  { name: "read_probe", description: "Read", input_schema: { type: "object", properties: {} } },
  { name: "write_probe", description: "Write", input_schema: { type: "object", properties: { value: { type: "string" } }, required: ["value"] } },
];
class ToolAdapter extends MockAdapter {
  contexts: StartSessionContext[] = [];
  closes: string[] = [];
  contextsById = new Map<string, StartSessionContext>();
  parallel = false;
  holdText = false;
  confirmNative = true;
  turns: SendTurnInput[] = [];
  received: unknown[] = [];
  override async startSession(ctx: StartSessionContext) { this.contexts.push(ctx); const id = String(this.contexts.length); this.contextsById.set(id, ctx); return id; }
  async relay(ctx: StartSessionContext, name: string, input: object) {
    const res = await fetch(`${ctx.mcp.env.AGENT_GATEWAY_RELAY_URL}/internal/tool`, { method: "POST", headers: {
      "content-type": "application/json", authorization: `Bearer ${ctx.mcp.env.AGENT_GATEWAY_RELAY_TOKEN}` },
      body: JSON.stringify({ session: ctx.sessionId, name, arguments: input }), signal: ctx.signal });
    if (!res.ok) throw new Error(`Relay ${res.status}`);
    const result = await res.json() as { content: Array<{ text: string }>; isError: boolean };
    this.received.push(result);
    return result;
  }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    this.turns.push(input);
    const ctx = this.contextsById.get(input.providerSessionId)!;
    const requestId = input.requestId;
    if (this.confirmNative) yield { type: "session_info", provider: this.id, providerSessionId: `native-${input.providerSessionId}`, requestId };
    yield { type: "message_delta", requestId, text: "Reading." };
    yield { type: "message_completed", requestId, text: "Reading." };
    if (this.holdText) {
      await new Promise<void>(resolve => input.signal.addEventListener("abort", () => resolve(), { once: true }));
      yield { type: "turn_finished", requestId, reason: "cancelled" };
      return;
    }
    const a = this.relay(ctx, "read_probe", {});
    const b = this.parallel ? this.relay(ctx, "read_probe", {}) : undefined;
    const result = await a; if (b) await b;
    const value = result.content[0]!.text;
    await this.relay(ctx, "write_probe", { value });
    yield { type: "message_delta", requestId, text: "DONE " };
    yield { type: "message_delta", requestId, text: value };
    yield { type: "message_completed", requestId, text: `DONE ${value}` };
    yield { type: "usage_update", requestId, usage: { inputTokens: 15, outputTokens: 9, totalTokens: 24, cachedInputTokens: 5 } };
    yield { type: "turn_finished", requestId, reason: "completed" };
  }
  override async closeSession(id: string) { this.closes.push(id); this.contextsById.delete(id); }
}
const servers: GatewayServer[] = [];
const directories: string[] = [];
afterEach(async () => { await Promise.all(servers.splice(0).map(s => s.close())); for (const directory of directories.splice(0)) fs.rmSync(directory, { recursive: true, force: true }); });
async function start(overrides: object = {}, conversationFile?: string, authToken = token) {
  const adapter = new ToolAdapter({ id: "copilot" });
  const config = { ...ConfigSchema.parse({ warmupOnStart: false, ...overrides }), port: 0 };
  const logs: object[] = [];
  const server = new GatewayServer({ config, token: authToken, tls: null, conversationFile, registry: new ProviderRegistry().register(adapter), log: entry => logs.push(entry) });
  servers.push(server); const url = await server.listen();
  const messages: ChatMessage[] = [{ role: "user", content: "read then write" }];
  const post = (body: object = {}, auth = authToken) => fetch(`${url}/v1/messages`, { method: "POST", headers: { "x-api-key": auth, "content-type": "application/json", origin: "https://pivot.claude.ai" },
    body: JSON.stringify({ model: "copilot/default", max_tokens: 1024, messages, tools, ...body }) });
  return { adapter, server, url, messages, post, logs };
}
function results(message: ApiMessage, value = "SENTINEL"): ChatMessage[] {
  return [{ role: "assistant", content: message.content }, { role: "user", content: message.content.filter(b => b.type === "tool_use")
    .map(b => ({ type: "tool_result", tool_use_id: b.id, content: value })) }];
}
describe("Anthropic-compatible CLI gateway", () => {
  it("routes explicit providers and aliases without silent provider fallback", () => {
    const config = ConfigSchema.parse({ aliases: { fast: { provider: "copilot", model: "small" } } });
    expect(routeModel("copilot/some/native/model", config).model).toBe("some/native/model");
    expect(routeModel("anthropic/copilot/default", config).model).toBeNull();
    expect(routeModel("copilot/default [Claude API]", config)).toMatchObject({ provider: "copilot", model: null });
    expect(routeModel("fast", config).model).toBe("small");
    // Bare claude-* names never fall back silently; only the login probe or an explicit opt-in may.
    expect(() => routeModel("claude-haiku-4-5-20251001", config)).toThrow("provider/model");
    expect(routeModel("claude-haiku-4-5-20251001", config, true).provider).toBe("copilot");
    expect(routeModel("claude-opus-4-7", ConfigSchema.parse({ advertiseCompatibilityModels: true })).provider).toBe("copilot");
    expect(() => routeModel("codex/default", config)).toThrow("not enabled");
    expect(() => routeModel("unknown/default", config)).toThrow("Unknown provider");
  });
  it("routes friendly model IDs precisely and rejects ambiguous names", () => {
    const config = ConfigSchema.parse({ enabledProviders: ["copilot", "codex"], modelNames: { "copilot/default": "Copilot — Default", "codex/gpt-6.1-sol": "Codex — GPT-6.1-Sol" } });
    expect(routeModel("Codex — GPT-6.1-Sol (Claude API)", config)).toMatchObject({ provider: "codex", model: "gpt-6.1-sol" });
    expect(routeModel("Copilot — Default (Claude API)", config)).toMatchObject({ provider: "copilot", model: null });
    expect(routeModel("codex/gpt-6.1-sol [Claude API]", config).model).toBe("gpt-6.1-sol");
    expect(() => ConfigSchema.parse({ modelNames: { "copilot/default": "Same", "codex/default": "Same" } })).toThrow("unique");
  });
  it("rejects unsupported sampling, server tools and image input explicitly", () => {
    const base = { model: "copilot/default", max_tokens: 32, messages: [{ role: "user", content: "hi" }] };
    expect(() => parseRequest({ ...base, temperature: 0 })).toThrow("Sampling");
    expect(() => parseRequest({ ...base, tools: [{ type: "web_search_20250305", name: "web_search" }] })).toThrow("input_schema");
    expect(() => parseRequest({ ...base, messages: [{ role: "user", content: [{ type: "image", source: {} }] }] })).toThrow("inline base64");
    expect(() => parseRequest({ ...base, messages: [{ role: "user", content: [{ type: "document", source: {} }] }] })).toThrow("Unsupported content block");
  });
  it("authenticates and returns CORS on errors and preflight", async () => {
    const h = await start();
    const pre = await fetch(`${h.url}/v1/messages`, { method: "OPTIONS", headers: { origin: "https://pivot.claude.ai", "access-control-request-headers": "x-api-key,authorization" } });
    expect(pre.status).toBe(204); expect(pre.headers.get("access-control-allow-private-network")).toBe("true");
    const denied = await h.post({}, "bad-token"); expect(denied.status).toBe(401); expect(denied.headers.get("access-control-allow-origin")).toBe("https://pivot.claude.ai");
    const cross = await fetch(`${h.url}/health`, { headers: { origin: "https://unrelated.example" } }); expect(cross.status).toBe(403);
    const internal = await fetch(`${h.url}/internal/tools`, { method: "POST", headers: { authorization: `Bearer ${token}` } }); expect(internal.status).toBe(401);
  });
  it("returns an HTTP 413 instead of resetting an oversized upload", async () => {
    const h = await start({ maxBodyBytes: 1024 });
    const res = await h.post({ messages: [{ role: "user", content: "x".repeat(4096) }] });
    expect(res.status).toBe(413);
    expect((await res.json()).error.type).toBe("request_too_large");
    expect(h.adapter.contexts).toHaveLength(0);
  });
  it("stops an active runtime when a streaming client disconnects", async () => {
    const h = await start();
    const controller = new AbortController();
    const res = await fetch(`${h.url}/v1/messages`, { method: "POST", signal: controller.signal, headers: { "x-api-key": token, "content-type": "application/json" },
      body: JSON.stringify({ model: "copilot/default", max_tokens: 100, messages: h.messages, tools, stream: true }) });
    await res.body!.getReader().read();
    controller.abort();
    await new Promise(resolve => setTimeout(resolve, 50));
    expect(h.adapter.contexts[0]?.signal.aborted).toBe(true);
    expect(h.adapter.closes).toEqual(["1"]);
  });
  it("advertises Claude-client aliases without changing their actual provider", async () => {
    const h = await start();
    const list = await (await fetch(`${h.url}/v1/models`, { headers: { "x-api-key": token } })).json();
    expect(list.data.some((m: { id: string }) => m.id === "copilot/default [Claude API]")).toBe(true);
    expect(list.data.some((m: { id: string }) => m.id === "claude-opus-4-7")).toBe(false);
  });
  it("advertises friendly names once while accepting technical model IDs", async () => {
    const h = await start({ modelNames: { "copilot/default": "Copilot — Default" } });
    const list = await (await fetch(`${h.url}/v1/models`, { headers: { "x-api-key": token } })).json();
    expect(list.data.map((m: { id: string }) => m.id)).toEqual(["Copilot — Default (Claude API)"]);
    const response = await h.post({ model: "Copilot — Default (Claude API)" });
    expect(response.status).toBe(200);
    expect(response.headers.get("x-agent-gateway-provider")).toBe("copilot");
    expect(response.headers.get("x-agent-gateway-model")).toBe("default");
  });
  it("uses a real backend result for cached login probes, never caches user messages", async () => {
    class ProbeAdapter extends MockAdapter {
      calls = 0;
      override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
        this.calls++;
        yield { type: "message_delta", requestId: input.requestId, text: "REAL_BACKEND_RESPONSE" };
        yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
      }
    }
    const adapter = new ProbeAdapter({ id: "copilot" });
    const config = { ...ConfigSchema.parse({ warmupOnStart: true }), port: 0 };
    const server = new GatewayServer({ config, token, tls: null, registry: new ProviderRegistry().register(adapter) });
    servers.push(server); const url = await server.listen();
    expect(adapter.calls).toBe(1);
    const probe = { model: "claude-opus-4-7", max_tokens: 1, messages: [{ role: "user", content: "x" }] };
    const post = (body: object) => fetch(`${url}/v1/messages`, { method: "POST", headers: { "x-api-key": token, "content-type": "application/json" }, body: JSON.stringify(body) });
    const cached = await post(probe);
    expect(cached.headers.get("x-agent-gateway-cache")).toBe("verification-probe");
    expect((await cached.json()).content[0].text).toBe("REAL_BACKEND_RESPONSE");
    expect(adapter.calls).toBe(1);
    const fresh = await post({ ...probe, model: "copilot/default", messages: [{ role: "user", content: "different" }] });
    expect(fresh.headers.get("x-agent-gateway-cache")).toBeNull();
    expect(adapter.calls).toBe(2);
    expect((await post({ ...probe, messages: [{ role: "user", content: "different" }] })).status).toBe(400);
    expect(adapter.calls).toBe(2);
  });
  it("round-trips client tools through one runtime and reports final usage once", async () => {
    const h = await start();
    const first = await (await h.post()).json() as ApiMessage;
    expect(first.stop_reason).toBe("tool_use"); expect(first.content).toContainEqual({ type: "text", text: "Reading." });
    expect(first.content.find(b => b.type === "tool_use")?.name).toBe("read_probe");
    h.messages.push(...results(first));
    const second = await (await h.post()).json() as ApiMessage;
    expect(second.content.find(b => b.type === "tool_use")).toMatchObject({ name: "write_probe", input: { value: "SENTINEL" } });
    h.messages.push(...results(second, "written"));
    const final = await (await h.post()).json() as ApiMessage;
    expect(final.stop_reason).toBe("end_turn"); expect(final.content).toEqual([{ type: "text", text: "DONE SENTINEL" }]);
    expect(final.usage).toMatchObject({ input_tokens: 10, output_tokens: 9, cache_read_input_tokens: 5 });
    expect(h.adapter.contexts).toHaveLength(1);
    expect(h.adapter.contexts[0]?.mcp.toolNames).toEqual(["read_probe", "write_probe"]);
  });
  async function complete(h: Awaited<ReturnType<typeof start>>, options: object = {}) {
    const read = await (await h.post(options)).json() as ApiMessage;
    h.messages.push(...results(read));
    const write = await (await h.post(options)).json() as ApiMessage;
    h.messages.push(...results(write, "written"));
    const final = await (await h.post(options)).json() as ApiMessage;
    expect(final.stop_reason).toBe("end_turn");
    h.messages.push({ role: "assistant", content: final.content });
  }
  const metadata = (id = "chat-123") => ({ metadata: { user_id: JSON.stringify({ session_id: id, user: "private-account", device_id: "private-device" }) } });
  function storage() { const directory = fs.mkdtempSync(path.join(os.tmpdir(), "gateway-conversations-")); directories.push(directory); return path.join(directory, "conversations.json"); }
  it("selects by client conversation ID and keeps identical chats isolated", async () => {
    const h = await start(); await complete(h, metadata("chat-a"));
    const sameHistory = structuredClone(h.messages);
    // Another ID with identical history must not attach to chat-a's worker.
    h.messages.push({ role: "user", content: "continue" });
    await complete(h, metadata("chat-b"));
    expect(h.adapter.contexts).toHaveLength(2);
    h.messages.splice(0, h.messages.length, ...sameHistory, { role: "user", content: "continue chat-a" });
    await complete(h, metadata("chat-a"));
    expect(h.adapter.turns.map(turn => turn.providerSessionId)).toEqual(["1", "2", "1"]);
    expect(h.logs).toContainEqual(expect.objectContaining({ phase: "start", match: "conversation_id", identitySource: "metadata.user_id.session_id" }));
  });
  it("persists native identity across gateway restart and sends only appended input", async () => {
    const file = storage();
    const first = await start({}, file); await complete(first, metadata()); await first.server.close();
    const saved = fs.readFileSync(file, "utf8");
    expect(saved).toContain('"nativeId":"native-1"');
    for (const secret of ["private-account", "private-device", "chat-123", "SENTINEL", "read then write"]) expect(saved).not.toContain(secret);
    const second = await start({}, file);
    second.messages.splice(0, second.messages.length, ...first.messages, { role: "user", content: "after gateway restart" });
    await complete(second, metadata());
    expect(second.adapter.contexts[0]!.resumeSessionId).toBe("native-1");
    expect(second.adapter.turns[0]!.prompt).toContain("after gateway restart");
    expect(second.adapter.turns[0]!.prompt).not.toContain("read then write");
    expect(second.logs).toContainEqual(expect.objectContaining({ phase: "start", match: "persisted_conversation_id" }));
  });
  it("binds persisted conversation IDs to the authenticated token", async () => {
    const file = storage(); const first = await start({}, file); await complete(first, metadata()); await first.server.close();
    const other = await start({}, file, "different-gateway-token-for-test");
    other.messages.splice(0, other.messages.length, ...first.messages, { role: "user", content: "another owner" });
    await other.post(metadata());
    expect(other.adapter.contexts[0]!.resumeSessionId).toBeUndefined();
  });
  it("rejects changed IDs on tool results and simultaneous turns for the same chat", async () => {
    const h = await start(); const first = await (await h.post(metadata())).json() as ApiMessage;
    expect((await h.post(metadata())).status).toBe(409);
    h.messages.push(...results(first));
    expect((await h.post(metadata("another-chat"))).status).toBe(409);
    expect((await h.post(metadata())).status).toBe(200);
  });
  it("refuses modified or repeated history for an explicit ID instead of creating a worker", async () => {
    const h = await start(); await complete(h, metadata());
    expect((await h.post(metadata())).status).toBe(409);
    h.messages[0] = { role: "user", content: "edited old message" };
    h.messages.push({ role: "user", content: "next" });
    expect((await h.post(metadata())).status).toBe(409);
    expect(h.adapter.contexts).toHaveLength(1);
  });
  it("does not replay interrupted tool turns after restart", async () => {
    const file = storage(); const h = await start({}, file);
    const first = await (await h.post(metadata())).json() as ApiMessage;
    await h.server.close();
    const restarted = await start({}, file);
    expect((await restarted.post(metadata())).status).toBe(409);
    expect((await restarted.post({ ...metadata(), messages: [...h.messages, ...results(first)] })).status).toBe(409);
    expect(restarted.adapter.contexts).toHaveLength(0);
  });
  it("lets a chat retry after Stop when the turn never called a tool", async () => {
    const file = storage(); const h = await start({}, file);
    h.adapter.holdText = true;
    const controller = new AbortController();
    const stopped = fetch(`${h.url}/v1/messages`, { method: "POST", signal: controller.signal, headers: { "x-api-key": token, "content-type": "application/json" },
      body: JSON.stringify({ model: "copilot/default", max_tokens: 1024, messages: h.messages, tools, ...metadata() }) }).catch(() => undefined);
    await vi.waitFor(() => expect(h.adapter.turns).toHaveLength(1));
    controller.abort(); await stopped;
    await vi.waitFor(() => expect(h.adapter.closes).toHaveLength(1));
    h.adapter.holdText = false;
    // Same chat, same messages: the interrupted attempt left nothing to replay.
    expect((await h.post(metadata())).status).toBe(200);
  });
  it.each([false, true])("accepts Office context_snip history changes only with explicit identity (%s)", async explicit => {
    const h = await start();
    const relay = h.adapter.relay.bind(h.adapter);
    h.adapter.relay = (ctx, name, input) => relay(ctx, name === "read_probe" ? "context_snip" : name, input);
    const options = { ...(explicit ? metadata() : {}), tools: [{ ...tools[0]!, name: "context_snip" }, tools[1]!] };
    const first = await (await h.post(options)).json() as ApiMessage;
    expect(first.content.find(block => block.type === "tool_use")?.name).toBe("context_snip");
    h.messages[0] = { role: "user", content: "[Earlier context was snipped by the client]" };
    h.messages.push(...results(first));
    const response = await h.post(options);
    expect(response.status).toBe(explicit ? 200 : 409);
    expect(h.adapter.contexts).toHaveLength(1);
    if (explicit) {
      const second = await response.json() as ApiMessage;
      h.messages.push(...results(second, "written"));
      const final = await (await h.post(options)).json() as ApiMessage;
      expect(final.stop_reason).toBe("end_turn");
      h.messages.push({ role: "assistant", content: final.content }, { role: "user", content: "after snip" });
      await h.post(options);
      expect(h.adapter.contexts).toHaveLength(1);
      expect(h.logs).toContainEqual(expect.objectContaining({ phase: "tool_result", historyRebased: true }));
    }
  });
  it.each([true, false])("converts image tool results and records conversion failures without losing the conversation (%s)", async valid => {
    const h = await start();
    const png = gatewayProbeImage().toString("base64");
    const first = await (await h.post()).json() as ApiMessage;
    const next = results(first);
    (next[1]!.content as ContentBlock[])[0]!.content = [
      { type: "text", text: "IMAGE_SENTINEL" },
      { type: "image", source: valid ? { type: "base64", media_type: "image/png", data: png } : { type: "url", url: "https://example.invalid/image" } },
    ];
    h.messages.push(...next);
    const second = await (await h.post()).json() as ApiMessage;
    expect(h.adapter.received[0]).toMatchObject({ isError: !valid, content: valid
      ? [{ type: "text", text: "IMAGE_SENTINEL" }, { type: "image", mimeType: "image/png", data: png }]
      : [{ type: "text", text: "IMAGE_SENTINEL" }, { type: "text", text: expect.stringContaining("may already have executed") }] });
    expect(h.logs).toContainEqual(expect.objectContaining({ phase: "tool_result", clientErrorCount: 0, conversionErrorCount: valid ? 0 : 1, errorCount: valid ? 0 : 1 }));
    h.messages.push(...results(second, "written"));
    const final = await (await h.post()).json() as ApiMessage;
    h.messages.push({ role: "assistant", content: final.content }, { role: "user", content: "next turn" });
    expect((await h.post()).status).toBe(200);
    expect(h.adapter.contexts).toHaveLength(1);
  });
  it.each([true, false])("keeps the live worker when Office redacts an older image (during pending result: %s)", async duringPending => {
    const h = await start();
    const first = await (await h.post()).json() as ApiMessage;
    const next = results(first);
    const result = (next[1]!.content as ContentBlock[])[0]!;
    const text = { type: "text", text: "IMAGE_CONTEXT" };
    result.content = [text, { type: "image", source: { type: "base64", media_type: "image/png", data: gatewayProbeImage().toString("base64") } }, { type: "text", text: "[id:abc]" }];
    h.messages.push(...next);
    const second = await (await h.post()).json() as ApiMessage;
    const redact = () => { result.content = [text, { type: "text", text: "[image previously shown — redacted from history]" }, { type: "text", text: "[id:abc]" }]; };
    if (duringPending) redact();
    h.messages.push(...results(second, "written"));
    const final = await (await h.post()).json() as ApiMessage;
    expect(final.stop_reason).toBe("end_turn");
    if (!duringPending) redact();
    h.messages.push({ role: "assistant", content: final.content }, { role: "user", content: "after Office image redaction" });
    await complete(h);
    expect(h.adapter.contexts).toHaveLength(1);
    expect(h.adapter.turns[1]!.prompt).not.toContain("IMAGE_CONTEXT");
    expect(h.logs).toContainEqual(expect.objectContaining({ phase: "runtime_ready", runtime: "reused" }));
  });
  it("rejects excessive prompt text before creating a provider session", async () => {
    const h = await start({ maxPromptChars: 3000 });
    const response = await h.post({ messages: [{ role: "user", content: "x".repeat(4000) }] });
    expect(response.status).toBe(413);
    expect(h.adapter.contexts).toHaveLength(0);
    expect(h.adapter.turns).toHaveLength(0);
  });
  it("keeps the worker across completed turns and sends only appended messages", async () => {
    const h = await start();
    await complete(h);
    expect(h.adapter.closes).toEqual([]);
    h.messages.push({ role: "user", content: "second turn" });
    await complete(h);
    expect(h.adapter.contexts).toHaveLength(1);
    expect(h.adapter.turns.map(t => t.providerSessionId)).toEqual(["1", "1"]);
    expect(h.adapter.turns[0]!.requestId).not.toBe(h.adapter.turns[1]!.requestId);
    expect(h.adapter.turns[1]!.prompt).toContain('"second turn"');
    expect(h.adapter.turns[1]!.prompt).not.toContain('"read then write"');
    expect(h.adapter.contexts[0]!.signal.aborted).toBe(false);
    await h.server.close();
    expect(h.adapter.closes).toEqual(["1"]);
  });
  it("reuses equivalent tool ordering, default tool choice and system representations", async () => {
    const h = await start(); await complete(h);
    h.messages.push({ role: "user", content: "same configuration" });
    const response = await h.post({ tools: [...tools].reverse(), tool_choice: { type: "auto", disable_parallel_tool_use: false }, system: [] });
    expect(response.status).toBe(200);
    expect(h.adapter.contexts).toHaveLength(1);
    expect(h.adapter.closes).toEqual([]);
    h.messages.push(...results(await response.json() as ApiMessage));
    expect((await h.post()).status).toBe(200);
    expect(h.logs).toContainEqual(expect.objectContaining({ event: "gateway_turn", phase: "runtime_ready", runtime: "reused" }));
  });
  it("matches split text and equivalent tool result text without ignoring edits", async () => {
    const h = await start(); await complete(h);
    for (const message of h.messages) {
      const content = typeof message.content === "string" ? [{ type: "text", text: message.content }] : message.content;
      message.content = content.flatMap((block): ContentBlock[] => {
        if (block.type === "text") {
          const text = String(block.text), mid = Math.floor(text.length / 2);
          return [{ type: "text", text: text.slice(0, mid) }, { type: "text", text: text.slice(mid), cache_control: { type: "ephemeral" } }];
        }
        return [{ ...block, ...(block.type === "tool_result" ? { content: [{ type: "text", text: block.content }] } : {}) }];
      });
    }
    h.messages.push({ role: "user", content: "continue equivalent history" });
    await h.post();
    expect(h.adapter.contexts).toHaveLength(1);
    expect(h.adapter.turns[1]!.prompt).not.toContain("SENTINEL");
  });
  it("logs lifecycle and latency without conversation or tool payloads", async () => {
    const h = await start(); await complete(h);
    expect(h.logs).toContainEqual(expect.objectContaining({ event: "provider_session", action: "runtime_opened", resume: false, durationMs: expect.any(Number) }));
    for (const phase of ["start", "runtime_ready", "first_text", "tool_wait", "tool_result", "completed"]) {
      expect(h.logs).toContainEqual(expect.objectContaining({ event: "gateway_turn", phase, sessionId: expect.any(String), elapsedMs: expect.any(Number) }));
    }
    expect(JSON.stringify(h.logs)).not.toContain("SENTINEL");
    expect(JSON.stringify(h.logs)).not.toContain("read then write");
  });
  it("uses shared native resume when an idle runtime dies", async () => {
    const h = await start(); await complete(h);
    h.adapter.contexts[0]!.onRuntimeClosed!();
    h.messages.push({ role: "user", content: "after process exit" });
    await complete(h);
    expect(h.adapter.contexts).toHaveLength(2);
    expect(h.adapter.closes).toEqual(["1"]);
    expect(h.adapter.contexts[1]!.resumeSessionId).toBe("native-1");
    expect(h.adapter.turns[1]!.prompt).not.toContain('"read then write"');
  });
  it("reconstructs history when a replacement runtime has no confirmed native ID", async () => {
    const h = await start(); h.adapter.confirmNative = false; await complete(h);
    h.adapter.contexts[0]!.onRuntimeClosed!();
    h.messages.push({ role: "user", content: "recover without native identity" });
    await h.post();
    expect(h.adapter.contexts[1]!.resumeSessionId).toBeUndefined();
    expect(h.adapter.turns[1]!.prompt).toContain('"read then write"');
    expect(h.adapter.turns[1]!.prompt).toContain('"recover without native identity"');
  });
  it("reopens configuration changes through the shared manager with native resume", async () => {
    const h = await start(); await complete(h);
    h.messages.push({ role: "user", content: "use another model" });
    const response = await h.post({ model: "copilot/other", system: "new instructions" });
    expect(response.status).toBe(200);
    expect(h.adapter.closes).toEqual(["1"]);
    expect(h.adapter.contexts[1]).toMatchObject({ model: "other", systemPrompt: expect.stringMatching(/^new instructions\n\n<gateway_tool_contract>/), resumeSessionId: "native-1" });
  });
  it.each([true, false])("delivers the API tool contract in session instructions, with Office rules only for Office tools (%s)", async officeJs => {
    const h = await start();
    const registered = officeJs ? [...tools, { name: "execute_office_js", input_schema: { type: "object", properties: {} } }] : tools;
    expect((await h.post({ system: "Caller system instructions", tools: registered })).status).toBe(200);
    const ctx = h.adapter.contexts[0]!;
    expect(ctx.systemPrompt).toMatch(/^Caller system instructions\n\n<gateway_tool_contract>/);
    expect(ctx.systemPromptMode).toBe("replace");
    expect(ctx.systemPrompt).toContain("remote MCP tools");
    expect(ctx.systemPrompt?.includes("shapes.addLine(startLeft, startTop, endLeft, endTop, Excel.ConnectorType.straight)")).toBe(officeJs);
    expect(ctx.systemPrompt?.includes("Chart.getData() is not an Office.js chart API")).toBe(officeJs);
  });
  it("keeps separate chats isolated, including concurrent branches of the same history", async () => {
    const h = await start(); await complete(h);
    const branch = [...h.messages, { role: "user" as const, content: "branch" }];
    await h.post({ messages: branch });
    await h.post({ messages: branch });
    await h.post({ messages: [{ role: "user", content: "another chat" }] });
    expect(h.adapter.turns.map(t => t.providerSessionId)).toEqual(["1", "1", "2", "3"]);
    expect(h.adapter.turns[2]!.prompt).toContain('"read then write"');
    expect(h.adapter.turns[3]!.prompt).not.toContain('"read then write"');
  });
  it("reconstructs altered history instead of attaching it to a live conversation", async () => {
    const h = await start(); await complete(h);
    h.messages[0] = { role: "user", content: "edited previous message" };
    h.messages.push({ role: "user", content: "followup" });
    await h.post();
    expect(h.adapter.contexts).toHaveLength(2);
    expect(h.adapter.turns[1]!.prompt).toContain('"edited previous message"');
    expect(h.adapter.contexts[1]!.resumeSessionId).toBeUndefined();
  });
  it("resumes native identity after idle expiry and capacity eviction without replay", async () => {
    const h = await start({ maxSessions: 1 });
    h.server.options.config.idleTimeoutMs = 40;
    await complete(h);
    const original = [...h.messages];
    h.messages.splice(0, h.messages.length, { role: "user", content: "new chat" });
    await complete(h);
    h.messages.splice(0, h.messages.length, ...original);
    expect(h.adapter.contexts).toHaveLength(2);
    expect(h.adapter.closes).toEqual(["1"]);
    h.messages.push({ role: "user", content: "after capacity eviction" });
    await h.post();
    expect(h.adapter.contexts[2]!.resumeSessionId).toBe("native-1");
    expect(h.adapter.turns[2]!.prompt).not.toContain('"read then write"');
    // A second server checks idle TTL independently of pending-tool TTL.
    const idle = await start(); idle.server.options.config.idleTimeoutMs = 30;
    await complete(idle);
    await new Promise(resolve => setTimeout(resolve, 70));
    expect(idle.adapter.closes).toEqual(["1"]);
    idle.messages.push({ role: "user", content: "after eviction" });
    await idle.post();
    expect(idle.adapter.contexts[1]!.resumeSessionId).toBe("native-1");
    expect(idle.adapter.turns[1]!.prompt).not.toContain('"read then write"');
  });
  it("accepts completed client bootstrap tools as initial conversation history", async () => {
    const h = await start();
    const bootstrap = [
      { role: "user", content: "Read my selection" },
      { role: "assistant", content: [{ type: "tool_use", id: "toolu_office_bootstrap", name: "get_context", input: {} }] },
      { role: "user", content: [{ type: "tool_result", tool_use_id: "toolu_office_bootstrap", content: "A1 selected" }] },
    ];
    const response = await h.post({ messages: bootstrap });
    expect(response.status).toBe(200);
    expect((await response.json()).stop_reason).toBe("tool_use");
    expect(h.adapter.contexts).toHaveLength(1);
  });
  it("accepts tool error results and forwards the original error marker", async () => {
    const h = await start();
    const first = await (await h.post()).json() as ApiMessage;
    const next = results(first, "not found");
    (next[1]!.content as Array<Record<string, unknown>>)[0]!.is_error = true;
    h.messages.push(...next);
    const response = await (await h.post()).json() as ApiMessage;
    expect(response.content.find(b => b.type === "tool_use")).toMatchObject({ input: { value: "not found" } });
  });
  it("rejects mismatched history and route before resolving a pending tool", async () => {
    const h = await start();
    const first = await (await h.post()).json() as ApiMessage;
    const good = [...h.messages, ...results(first)];
    expect((await h.post({ messages: [...results(first)] })).status).toBe(409);
    expect((await h.post({ messages: good, model: "copilot/other" })).status).toBe(409);
    h.messages.push(...results(first));
    expect((await h.post()).status).toBe(200);
  });
  it("batches parallel tools and requires all results", async () => {
    const h = await start(); h.adapter.parallel = true;
    const first = await (await h.post()).json() as ApiMessage;
    expect(first.content.filter(b => b.type === "tool_use")).toHaveLength(2);
    const next = results(first);
    const partial = [...h.messages, next[0]!, { ...next[1]!, content: (next[1]!.content as object[]).slice(0, 1) }];
    expect((await h.post({ messages: partial })).status).toBe(400);
    h.messages.push(...next); expect((await h.post()).status).toBe(200);
  });
  it("emits valid SSE tool blocks and a single stream termination", async () => {
    const h = await start();
    const response = await h.post({ stream: true }); const text = await response.text();
    expect(response.headers.get("content-type")).toContain("text/event-stream");
    const events = text.split("\n").filter(l => l.startsWith("data: ")).map(l => JSON.parse(l.slice(6)));
    expect(events.filter(e => e.type === "message_stop")).toHaveLength(1);
    expect(events.find(e => e.type === "content_block_start" && e.content_block.type === "tool_use").content_block.name).toBe("read_probe");
    expect(events.find(e => e.type === "message_delta").delta.stop_reason).toBe("tool_use");
  });
  it("limits parallel sessions and never consumes a continuation twice", async () => {
    const h = await start({ maxSessions: 1 });
    const first = await (await h.post()).json() as ApiMessage;
    expect((await h.post()).status).toBe(429);
    h.messages.push(...results(first));
    expect((await h.post()).status).toBe(200);
    expect((await h.post()).status).toBe(409);
  });
  it("expires unanswered tools without replaying them", async () => {
    const h = await start(); h.server.options.config.continuationTimeoutMs = 30;
    const first = await (await h.post()).json() as ApiMessage;
    await new Promise(resolve => setTimeout(resolve, 70));
    h.messages.push(...results(first));
    expect((await h.post()).status).toBe(409);
    expect(h.adapter.contexts).toHaveLength(1);
  });
});
