import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput } from "../providers/src/types.js";
import { ReferenceFiles } from "../bridge/src/reference-files.js";
import { SkillStore } from "../bridge/src/skills.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;

afterEach(async () => {
  await bridge?.close();
  bridge = null;
});

async function startBridge() {
  const port = await getFreePort();
  const config = loadConfig({
    host: "127.0.0.1",
    port,
    tls: null,
    publicBaseUrl: `http://127.0.0.1:${port}`,
    logLevel: "error",
  });
  const registry = new ProviderRegistry().register(
    new MockAdapter({
      id: "claude",
      toolCalls: [{ tool: "get_selection", args: {} }],
      reply: "Edited A1.",
    }),
  );
  bridge = new Bridge(config, { usageFile: null, registry });
  const { url } = await bridge.listen();
  return { url, port };
}

async function handshake(pane: PaneClient) {
  pane.send({
    type: "hello",
    protocolVersion: 1,
    workbookConnectionId: "0123456789abcdef",
  });
  const hello = await pane.waitFor((m) => m.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("expected hello_ok");
  return hello;
}

describe("bridge end to end", () => {
  it("saves an AI-created skill through MCP and exposes it as a manageable user skill", async () => {
    const folder = fs.mkdtempSync(path.join(os.tmpdir(), "skill-save-api-"));
    try {
      const port = await getFreePort();
      const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
      bridge = new Bridge(config, { usageFile: null, skills: new SkillStore([folder]), skillUploadDir: folder,
        registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", toolCalls: [{ tool: "save_skill",
          args: { name: "custom-report", description: "User reporting rules", instructions: "Format the selected table." } }] })),
      });
      const { url } = await bridge.listen();
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      const hello = await handshake(pane);
      pane.send({ type: "set_provider", provider: "claude" });
      await pane.waitFor(message => message.type === "provider_status");
      pane.send({ type: "prompt", requestId: "save-test", prompt: "Create and save a reporting skill.", selection: null });
      expect(await pane.waitFor(message => message.type === "tool_finished")).toMatchObject({ tool: "save_skill", ok: true });
      await pane.waitFor(message => message.type === "turn_finished");
      expect(fs.readFileSync(path.join(folder, "custom-report", "SKILL.md"), "utf8")).toContain("Format the selected table.");
      expect(await (await fetch(`${url}/skills`)).json()).toMatchObject({ skills: [{ name: "custom-report" }] });
      const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
      expect(await (await fetch(`${url}/session/skills`, { headers })).json()).toMatchObject({ skills: [{ name: "custom-report", removable: true }] });
      expect((await fetch(`${url}/session/skills`, { method: "DELETE", headers, body: JSON.stringify({ name: "custom-report" }) })).status).toBe(200);
      await pane.close();
    } finally { fs.rmSync(folder, { recursive: true, force: true }); }
  });
  it("uploads user skills through an authenticated session and exposes them in the slash catalog", async () => {
    const folder = fs.mkdtempSync(path.join(os.tmpdir(), "skill-upload-api-"));
    try {
      const port = await getFreePort();
      const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
      bridge = new Bridge(config, { usageFile: null, skills: new SkillStore([folder]), skillUploadDir: folder, registry: new ProviderRegistry().register(new MockAdapter({ id: "copilot" })) });
      const { url } = await bridge.listen();
      expect((await fetch(`${url}/session/skills`, { method: "POST", body: "test" })).status).toBe(401);
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      const hello = await handshake(pane);
      const headers = { authorization: `Bearer ${hello.sessionToken}` };
      const upload = await fetch(`${url}/session/skills`, { method: "POST", headers: { ...headers, "x-file-name": "report.md" }, body: "---\nname: report\ndescription: Custom report\n---\nUse the active selection." });
      expect(upload.status).toBe(201);
      expect(await (await fetch(`${url}/skills`)).json()).toMatchObject({ skills: [{ name: "report" }] });
      expect(await (await fetch(`${url}/session/skills`, { headers })).json()).toMatchObject({ skills: [{ name: "report", removable: true }] });
      const duplicate = await fetch(`${url}/session/skills`, { method: "POST", headers: { ...headers, "x-file-name": "report.md" }, body: "---\nname: report\n---\nDifferent instructions." });
      expect(duplicate.status).toBe(400);
      expect((await fetch(`${url}/session/skills`, { method: "DELETE", headers: { ...headers, "content-type": "application/json" }, body: JSON.stringify({ name: "report" }) })).status).toBe(200);
      expect(await (await fetch(`${url}/skills`)).json()).toMatchObject({ skills: [] });
      await pane.close();
    } finally { fs.rmSync(folder, { recursive: true, force: true }); }
  });
  it("reports provider authentication errors to the model picker", async () => {
    const { url } = await startBridge();
    class SignedOutProvider extends MockAdapter {
      async listModels(): Promise<never> { throw new Error("Authentication required"); }
    }
    bridge!.providers.register(new SignedOutProvider({ id: "copilot" }));
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const hello = await handshake(pane);
    const response = await fetch(`${url}/session/models?provider=copilot`, { headers: { authorization: `Bearer ${hello.sessionToken}` } });
    expect(response.status).toBe(401);
    expect(await response.json()).toMatchObject({ code: "login_required" });
    await pane.close();
  });
  it("requires a connected workbook token for device-code login and isolates its status", async () => {
    const { url } = await startBridge();
    for (const method of ["GET", "POST", "DELETE"]) {
      expect((await fetch(`${url}/session/provider-login`, { method })).status).toBe(401);
    }
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const hello = await handshake(pane);
    const headers = { authorization: `Bearer ${hello.sessionToken}` };
    const response = await fetch(`${url}/session/provider-login`, { headers });
    expect(response.headers.get("cache-control")).toBe("no-store");
    expect(await response.json()).toMatchObject({ state: "idle", userCode: null, verificationUri: null });
    expect((await fetch(`${url}/session/provider-login`, { method: "PATCH", headers })).status).toBe(405);
    expect(await (await fetch(`${url}/session/provider-login`, { method: "DELETE", headers })).json()).toMatchObject({ state: "idle" });
    await pane.close();
  });
  it("passes two selected paths to the CLI without copying either file", async () => {
    const folder = fs.mkdtempSync(path.join(os.tmpdir(), "selected-paths-"));
    const first = path.join(folder, "first.md");
    const second = path.join(folder, "second.zip");
    fs.writeFileSync(first, "# First");
    fs.writeFileSync(second, "PK sample");
    let receivedPrompt = "";
    class CaptureAdapter extends MockAdapter {
      override async *sendTurn(input: SendTurnInput) {
        receivedPrompt = input.prompt;
        yield* super.sendTurn(input);
      }
    }
    try {
      const port = await getFreePort();
      const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
      bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new CaptureAdapter({ id: "claude" })), references: new ReferenceFiles([folder]) });
      const { url } = await bridge.listen();
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      const hello = await handshake(pane);
      pane.send({ type: "set_provider", provider: "claude" });
      await pane.waitFor((message) => message.type === "provider_status");
      pane.send({ type: "prompt", requestId: "two-files", prompt: "Compare them", selection: null, referencePaths: [first, second] });
      await pane.waitFor((message) => message.type === "turn_finished" && message.requestId === "two-files");
      expect(receivedPrompt).toContain(JSON.stringify({ name: "first.md", path: first, kind: "reference" }));
      expect(receivedPrompt).toContain(JSON.stringify({ name: "second.zip", path: second, kind: "reference" }));
      expect(bridge.sessions.getByToken(hello.sessionToken)!.attachments.size).toBe(0);
      await pane.close();
    } finally { fs.rmSync(folder, { recursive: true, force: true }); }
  });
  it("passes a file selected in the task pane to the CLI as a local attachment", async () => {
    let receivedPrompt = "";
    class CaptureAdapter extends MockAdapter {
      override async *sendTurn(input: SendTurnInput) {
        receivedPrompt = input.prompt;
        yield* super.sendTurn(input);
      }
    }
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new CaptureAdapter({ id: "claude" })) });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const hello = await handshake(pane);
    const uploaded = await fetch(`${url}/session/attachment`, {
      method: "POST",
      headers: { authorization: `Bearer ${hello.sessionToken}`, "x-file-name": "notes.md" },
      body: "# Notes",
    });
    expect(uploaded.status).toBe(200);
    const file = await uploaded.json() as { id: string; path: string };
    pane.send({ type: "set_provider", provider: "claude" });
    await pane.waitFor((message) => message.type === "provider_status");
    pane.send({ type: "prompt", requestId: "local-file", prompt: "Read this", selection: null, attachmentIds: [file.id] });
    await pane.waitFor((message) => message.type === "turn_finished" && message.requestId === "local-file");
    expect(receivedPrompt).toContain(JSON.stringify({ name: "notes.md", path: file.path, kind: "file" }));
    await pane.close();
  });
  it("exposes health over HTTP", async () => {
    const { url } = await startBridge();
    const res = await fetch(`${url}/health`);
    expect(res.status).toBe(200);
    const body = (await res.json()) as { status: string; protocolVersion: number };
    expect(body.status).toBe("ok");
    expect(body.protocolVersion).toBe(1);
  });

  it("serves the provider model catalog to the connected pane", async () => {
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    class CatalogAdapter extends MockAdapter {
      async listModels() {
        return { models: [{ id: "current-model", name: "Current Model" }], defaultModelId: "current-model" };
      }
    }
    bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new CatalogAdapter({ id: "claude" })) });
    const { url } = await bridge.listen();
    const unauthenticated = await fetch(`${url}/session/models?provider=claude`);
    expect(unauthenticated.status).toBe(401);
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const hello = await handshake(pane);
    const response = await fetch(`${url}/session/models?provider=claude`, {
      headers: { authorization: `Bearer ${hello.sessionToken}` },
    });
    expect(response.status).toBe(200);
    expect(await response.json()).toEqual({ models: [{ id: "current-model", name: "Current Model" }], defaultModelId: "current-model" });
    await pane.close();
  });

  it("keeps a user-selected image in the connected workbook session", async () => {
    const { url } = await startBridge();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const hello = await handshake(pane);
    const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]);
    const uploaded = await fetch(`${url}/session/attachment`, {
      method: "POST",
      headers: { authorization: `Bearer ${hello.sessionToken}`, "x-file-name": "photo.png" },
      body: png,
    });
    expect(uploaded.status).toBe(200);
    const file = await uploaded.json() as { id: string; path: string; kind: string };
    expect(file.kind).toBe("image");
    const session = bridge!.sessions.getByToken(hello.sessionToken)!;
    expect(await bridge!.gateway.invoke({ sessionId: session.id, tool: "read_attached_image", args: { path: file.path } })).toMatchObject({ ok: true });
    await pane.close();
    expect(session.attachments.size).toBe(0);
  });

  it("manages skill templates and exposes them to skill tools", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "skill-templates-"));
    fs.mkdirSync(path.join(root, "skills", "sample"), { recursive: true });
    fs.writeFileSync(path.join(root, "skills", "sample", "SKILL.md"), "---\nname: sample\ndescription: Sample skill\n---\nFollow the example.");
    const skills = new SkillStore([path.join(root, "skills")], path.join(root, "templates"));
    try {
      const port = await getFreePort();
      const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
      bridge = new Bridge(config, { usageFile: null, skills, registry: new ProviderRegistry() });
      const { url } = await bridge.listen();
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      const hello = await handshake(pane);
      const auth = { authorization: `Bearer ${hello.sessionToken}` };
      const upload = await fetch(`${url}/session/skill-templates`, { method: "POST", headers: { ...auth, "x-skill-name": "sample", "x-file-name": "example.md" }, body: "Example layout" });
      expect(upload.status).toBe(200);
      const template = (await upload.json() as { template: { path: string } }).template;
      expect(template.path).toBe("skill-template://sample/example.md");
      const listed = await fetch(`${url}/session/skill-templates?skill=sample`, { headers: auth });
      expect((await listed.json() as { templates: { name: string }[] }).templates[0]!.name).toBe("example.md");
      const session = bridge.sessions.getByToken(hello.sessionToken)!;
      expect(await bridge.gateway.invoke({ sessionId: session.id, tool: "read_skill", args: { name: "sample", file: "templates/example.md" } }))
        .toMatchObject({ ok: true, data: { content: "Example layout" } });
      const deleted = await fetch(`${url}/session/skill-templates`, { method: "DELETE", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ skill: "sample", name: "example.md" }) });
      expect(deleted.status).toBe(200);
      expect(skills.listTemplates("sample")).toEqual([]);
      await pane.close();
    } finally { fs.rmSync(root, { recursive: true, force: true }); }
  });

  it("registers the open workbook folder and browses an added workspace", async () => {
    const folder = fs.mkdtempSync(path.join(os.tmpdir(), "workspace-"));
    const workbook = path.join(folder, "open.xlsx");
    const image = path.join(folder, "screen.png");
    const notes = path.join(folder, "notes.md");
    const extra = path.join(folder, "extra");
    const configFile = path.join(folder, "config.json");
    fs.mkdirSync(extra);
    fs.writeFileSync(workbook, "PK workbook");
    fs.writeFileSync(image, Buffer.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]));
    fs.writeFileSync(notes, "# Original file");
    try {
      // This test writes only to its temporary config, never the user's local settings.
      const port = await getFreePort();
      const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, referenceDirs: [], logLevel: "error" });
      bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry(), workspaceConfigPath: configFile });
      const live = await bridge.listen();
      const pane = await PaneClient.connect(`${live.url.replace("http", "ws")}/ws`);
      const hello = await handshake(pane);
      const auth = { authorization: `Bearer ${hello.sessionToken}` };
      const roots = await fetch(`${live.url}/session/workspaces?documentUrl=${encodeURIComponent(`file://${workbook}`)}`, { headers: auth });
      expect((await roots.json() as { roots: { path: string; isDefault: boolean }[] }).roots).toContainEqual(expect.objectContaining({ path: folder, isDefault: true }));
      const entries = await fetch(`${live.url}/session/workspaces/entries?path=${encodeURIComponent(folder)}`, { headers: auth });
      expect((await entries.json() as { entries: { name: string }[] }).entries.map((entry) => entry.name)).toEqual(["extra", "notes.md", "open.xlsx", "screen.png"]);
      const attached = await fetch(`${live.url}/session/workspaces/attach`, { method: "POST", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ path: image }) });
      expect(await attached.json()).toMatchObject({ path: image, kind: "file" });
      const attachedNotes = await fetch(`${live.url}/session/workspaces/attach`, { method: "POST", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ path: notes }) });
      expect(await attachedNotes.json()).toMatchObject({ path: notes, kind: "file" });
      const session = bridge.sessions.getByToken(hello.sessionToken)!;
      expect(session.attachments.size).toBe(0);
      expect(await bridge.gateway.invoke({ sessionId: session.id, tool: "read_attached_file", args: { path: notes } }))
        .toMatchObject({ ok: true, data: { content: "# Original file" } });
      expect(await bridge.gateway.invoke({ sessionId: session.id, tool: "read_attached_image", args: { path: image } }))
        .toMatchObject({ ok: true, data: { image: { mimeType: "image/png" } } });
      const outside = await fetch(`${live.url}/session/workspaces/entries?path=${encodeURIComponent(os.tmpdir())}`, { headers: auth });
      expect(outside.status).toBe(400);
      const added = await fetch(`${live.url}/session/workspaces`, { method: "POST", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ path: extra }) });
      expect(added.status).toBe(200);
      expect(JSON.parse(fs.readFileSync(configFile, "utf8")).referenceDirs).toEqual([extra]);
      const listed = await fetch(`${live.url}/session/workspaces`, { headers: auth });
      expect((await listed.json() as { roots: { path: string; removable: boolean }[] }).roots).toContainEqual(expect.objectContaining({ path: extra, removable: true }));
      const remove = (target: string) => fetch(`${live.url}/session/workspaces`, { method: "DELETE", headers: { ...auth, "content-type": "application/json" }, body: JSON.stringify({ path: target }) });
      expect((await remove(folder)).status).toBe(400);
      expect((await remove(extra)).status).toBe(200);
      expect(JSON.parse(fs.readFileSync(configFile, "utf8")).referenceDirs).toEqual([]);
      await pane.close();
    } finally {
      fs.rmSync(folder, { recursive: true, force: true });
    }
  });

  it("runs a prompt, routes a tool call to the pane and streams the result", async () => {
    const { url } = await startBridge();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    await handshake(pane);

    pane.send({ type: "set_provider", provider: "claude" });
    const status = await pane.waitFor((m) => m.type === "provider_status");
    expect(status).toMatchObject({
      type: "provider_status",
      provider: "claude",
      detection: { installed: true, status: "ready" },
    });

    pane.send({ type: "prompt", requestId: "req-1", prompt: "format the selection", selection: null });

    const toolRequest = await pane.waitFor((m) => m.type === "tool_request");
    if (toolRequest.type !== "tool_request") throw new Error("expected tool_request");
    expect(toolRequest.tool).toBe("get_selection");
    expect(toolRequest.requestId).toBe("req-1");

    pane.send({
      type: "tool_result",
      callId: toolRequest.callId,
      result: { ok: true, data: { address: "A1", rowCount: 1, columnCount: 1 } },
    });

    const finished = await pane.waitFor((m) => m.type === "turn_finished");
    expect(finished).toMatchObject({ type: "turn_finished", requestId: "req-1", reason: "completed" });

    const delta = await pane.waitFor((m) => m.type === "message_delta");
    expect(delta).toMatchObject({ type: "message_delta", text: "Edited A1." });

    const toolFinished = await pane.waitFor((m) => m.type === "tool_finished");
    expect(toolFinished).toMatchObject({ ok: true, tool: "get_selection" });

    await pane.close();
  });

  it("rejects a prompt when no provider is selected", async () => {
    const { url } = await startBridge();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    await handshake(pane);

    pane.send({ type: "prompt", requestId: "req-2", prompt: "hi", selection: null });
    const error = await pane.waitFor((m) => m.type === "error");
    expect(error).toMatchObject({ type: "error", code: "no_provider" });

    await pane.close();
  });

  it("rejects a protocol mismatch", async () => {
    const { url } = await startBridge();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    pane.send({ type: "hello", protocolVersion: 999, workbookConnectionId: "0123456789abcdef" });
    const rejected = await pane.waitFor((m) => m.type === "hello_error");
    expect(rejected).toMatchObject({ type: "hello_error", code: "protocol_mismatch" });
    await pane.close();
  });
});

describe("continued chats", () => {
  it("defaults to medium effort and restarts with the selected effort for the next turn", async () => {
    const { url } = await startBridge();
    const started: string[] = [];
    const received: string[] = [];
    class EffortAdapter extends MockAdapter {
      override async startSession(context: Parameters<MockAdapter["startSession"]>[0]) {
        started.push(context.effort ?? "missing");
        return super.startSession(context);
      }
      override async *sendTurn(input: SendTurnInput) {
        received.push(input.effort ?? "missing");
        yield* super.sendTurn(input);
      }
    }
    bridge!.providers.register(new EffortAdapter({ id: "claude" }));
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    await handshake(pane);
    pane.send({ type: "set_provider", provider: "claude" });
    await pane.waitFor(message => message.type === "provider_status");
    for (const [requestId, effort] of [["effort-default", undefined], ["effort-high", "high"], ["effort-high-again", "high"]] as const) {
      pane.send({ type: "prompt", requestId, prompt: "Hello", selection: null, ...(effort ? { effort } : {}) });
      await pane.waitFor(message => message.type === "turn_finished" && message.requestId === requestId);
    }
    expect(started).toEqual(["medium", "high"]);
    expect(received).toEqual(["medium", "high", "high"]);
    await pane.close();
  });
  it("starts a new provider session when the selected model changes", async () => {
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    const started: Array<string | null> = [];
    const closed: string[] = [];
    class Recording extends MockAdapter {
      override async startSession(ctx: Parameters<MockAdapter["startSession"]>[0]) {
        started.push(ctx.model ?? null);
        return super.startSession(ctx);
      }
      override async closeSession(id: string) {
        closed.push(id);
        return super.closeSession(id);
      }
    }
    bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new Recording({ id: "claude", reply: "ok" })) });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    await handshake(pane);
    pane.send({ type: "set_provider", provider: "claude" });
    await pane.waitFor((m) => m.type === "provider_status");

    for (const [index, model] of [null, null, "sonnet", "opus"].entries()) {
      const requestId = `model-${index}`;
      pane.send({ type: "prompt", requestId, prompt: "hello", selection: null, model });
      await pane.waitFor((m) => m.type === "turn_finished" && m.requestId === requestId);
    }
    expect(started).toEqual([null, "sonnet", "opus"]);
    expect(closed).toHaveLength(2);
    await pane.close();
  });

  it("prepends the replayed context to the prompt sent to the provider, after skill expansion", async () => {
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    const seen: string[] = [];
    class Recording extends MockAdapter {
      override async *sendTurn(input: Parameters<MockAdapter["sendTurn"]>[0]) {
        seen.push(input.prompt);
        yield* super.sendTurn(input);
      }
    }
    const registry = new ProviderRegistry().register(new Recording({ id: "claude", reply: "ok" }));
    bridge = new Bridge(config, { usageFile: null, registry });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    await handshake(pane);
    pane.send({ type: "set_provider", provider: "claude" });
    await pane.waitFor((m) => m.type === "provider_status");

    pane.send({ type: "prompt", requestId: "r1", prompt: "continue please", selection: null, context: "<previous_conversation>User: hi</previous_conversation>" });
    await pane.waitFor((m) => m.type === "turn_finished");
    expect(seen[0]).toContain("<previous_conversation>User: hi</previous_conversation>");
    expect(seen[0]?.endsWith("continue please")).toBe(true);

    pane.send({ type: "prompt", requestId: "r2", prompt: "no context here", selection: null });
    await pane.waitFor((m, i) => m.type === "turn_finished" && (m as { requestId: string }).requestId === "r2");
    expect(seen[1]?.endsWith("no context here")).toBe(true);
    await pane.close();
  });
});
