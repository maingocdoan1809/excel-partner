import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import type { ExternalSessionSummary, ImportedSession } from "../protocol/src/external-sessions.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { MockAdapter } from "../providers/src/mock.js";
import { PaneClient } from "./helpers.js";

let bridge: Bridge, pane: PaneClient, dir: string;
afterEach(async () => { await pane?.close(); await bridge?.close(); if (dir) fs.rmSync(dir, { recursive: true, force: true }); });

const summary: ExternalSessionSummary = { id: "native-1", title: "From the terminal", origin: "Terminal", updatedAt: 5 };
class Importing extends MockAdapter {
  override readonly capabilities = { ...MockAdapter.capabilities, importSessions: true };
  constructor() { super({ id: "copilot" }); }
  async listExternalSessions() { return [summary]; }
  async importExternalSession(id: string): Promise<ImportedSession> {
    if (id === "missing") throw new Error("That conversation was not found on this computer.");
    return { providerSessionId: id, title: "From the terminal", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "hello" }, { kind: "assistant", text: "hi" }] };
  }
}

async function fixture(adapter: MockAdapter) {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), "external-api-"));
  const config = loadConfig({ host: "127.0.0.1", port: 0, tls: null, configFile: path.join(dir, "config.json"), logLevel: "error" });
  bridge = new Bridge(config, { registry: new ProviderRegistry().register(adapter), usageFile: null, systemPromptFile: null, savedReferencesFile: null, updater: null });
  const { url } = await bridge.listen(); pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef", clientInfo: { host: "word" } } as never);
  const hello = await pane.waitFor(m => m.type === "hello_ok"); if (hello.type !== "hello_ok") throw new Error("hello");
  const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
  const list = (provider = "copilot") => fetch(`${url}/session/external?provider=${provider}`, { headers });
  const post = (body: unknown) => fetch(`${url}/session/external`, { method: "POST", headers, body: JSON.stringify(body) });
  const history = (query = "") => fetch(`${url}/session/history${query}`, { headers }).then(response => response.json() as Promise<any>);
  return { url, list, post, history };
}

describe("importing conversations from the provider's own tools", () => {
  it("needs a connected pane", async () => {
    const w = await fixture(new Importing());
    expect((await fetch(`${w.url}/session/external?provider=copilot`)).status).toBe(401);
  });

  it("lists what the provider offers, and refuses a provider that cannot import", async () => {
    const w = await fixture(new Importing());
    expect(await (await w.list()).json()).toEqual({ sessions: [summary] });
    expect((await w.list("claude")).status).toBe(400);
    expect((await w.list("nonsense")).status).toBe(400);
  });

  it("refuses when the provider does not declare it can import", async () => {
    const w = await fixture(new MockAdapter({ id: "copilot" }));
    expect((await w.list()).status).toBe(400);
    expect((await w.post({ provider: "copilot", id: "native-1" })).status).toBe(400);
  });

  it("saves the conversation in History under the pane's app, resumable by its native id, and keeps one copy however often it is imported", async () => {
    const w = await fixture(new Importing());
    const first = await (await w.post({ provider: "copilot", id: "native-1" })).json();
    expect(first).toEqual({ id: "copilot:native-1", existing: false });
    expect(await (await w.post({ provider: "copilot", id: "native-1" })).json()).toEqual({ id: "copilot:native-1", existing: true });
    const chat = await w.history("?view=chat&id=copilot%3Anative-1");
    expect(chat.session).toMatchObject({ id: "copilot:native-1", provider: "copilot", host: "word", providerSessionId: "native-1", identityVersion: 1, title: "From the terminal" });
    expect(chat.session.messages.map((message: { kind: string }) => message.kind)).toEqual(["user", "assistant"]);
    expect((await w.history()).total).toBe(1);
  });

  it("answers with the reason when the conversation cannot be read", async () => {
    const w = await fixture(new Importing());
    const response = await w.post({ provider: "copilot", id: "missing" });
    expect(response.status).toBe(409);
    expect(await response.json()).toEqual({ error: "That conversation was not found on this computer." });
    expect((await w.post({ provider: "copilot" })).status).toBe(400);
  });
});
