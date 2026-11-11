// @vitest-environment jsdom
import fs from "node:fs";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";
import { FakeDeck, installFakePowerPoint } from "./support/fake-powerpoint.js";

const html = fs.readFileSync(path.join(__dirname, "..", "addin", "taskpane", "index.html"), "utf8");
const g = globalThis as Record<string, any>;

class FakeSocket {
  static CONNECTING = 0; static OPEN = 1; static CLOSING = 2; static CLOSED = 3;
  static instances: FakeSocket[] = [];
  readyState = 0;
  sent: any[] = [];
  onopen: (() => void) | null = null;
  onmessage: ((event: { data: string }) => void) | null = null;
  onclose: (() => void) | null = null;
  onerror: (() => void) | null = null;
  constructor(public url: string) { FakeSocket.instances.push(this); queueMicrotask(() => { this.readyState = 1; this.onopen?.(); }); }
  send(text: string) { this.sent.push(JSON.parse(text)); }
  close() { this.readyState = 3; }
}

async function boot(host: "Excel" | "Word" | "PowerPoint", connected = false) {
  vi.resetModules();
  FakeSocket.instances = [];
  document.documentElement.innerHTML = html.replace(/<script[^>]*><\/script>/g, "");
  let ready: ((info: { host: string }) => void) | undefined;
  const office = g.Office ?? {};
  g.Office = {
    ...office,
    HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
    onReady: (callback: (info: { host: string }) => void) => { ready = callback; return Promise.resolve({ host }); },
    context: { ...office.context, requirements: office.context?.requirements ?? { isSetSupported: () => true }, host, platform: "PC", displayLanguage: "en-US", document: { ...office.context?.document, settings: { get: () => null, set: () => undefined, saveAsync: () => undefined }, addHandlerAsync: vi.fn(), removeHandlerAsync: vi.fn() }, ui: {} },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  };
  g.WebSocket = FakeSocket;
  g.ResizeObserver = class { observe() {} disconnect() {} unobserve() {} };
  g.matchMedia ??= () => ({ matches: false, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} });
  window.matchMedia ??= g.matchMedia;
  const saved = new Map<string, any>();
  let generation = 0;
  g.fetch = vi.fn(async (target: string, options: RequestInit = {}) => {
    const url = new URL(String(target), "https://localhost:8765");
    let data: unknown = {};
    if (url.pathname === "/session/history") {
      const input = options.body ? JSON.parse(String(options.body)) : {};
      if (options.method === "POST" && input.action === "import") for (const item of input.sessions) if (!saved.has(item.id)) saved.set(item.id, { ...item, host: host.toLowerCase(), revision: 1 });
      if (options.method === "POST" && input.action === "save") saved.set(input.session.id, { ...input.session, revision: (saved.get(input.session.id)?.revision ?? 0) + 1 });
      if (options.method === "DELETE") { const id = url.searchParams.get("id"); if (id) saved.delete(id); else { saved.clear(); generation++; } }
      const view = url.searchParams.get("view");
      if (view === "stats" || options.method === "DELETE") data = { count: saved.size, bytes: 100, generation };
      else if (view === "chat") { const item = saved.get(url.searchParams.get("id")!); data = { session: item, revision: item?.revision ?? 0, generation, totalMessages: item?.messages.length ?? 0, totalChanges: item?.changes.length ?? 0 }; }
      else if (input.action === "save") data = { revision: saved.get(input.session.id)?.revision, generation };
      else data = { items: [...saved.values()].map(item => ({ ...item, messageCount: item.messages.length, changeCount: item.changes.length })), total: saved.size, generation };
    }
    return new Response(JSON.stringify(data), { status: 200 });
  });
  await import("../addin/taskpane/main.js");
  expect(ready, "main.ts must register Office.onReady").toBeTypeOf("function");
  ready!({ host });
  await new Promise((resolve) => setTimeout(resolve, 20));
  const socket = FakeSocket.instances[0]!;
  if (connected) { socket.onmessage?.({ data: JSON.stringify({ type: "hello_ok", protocolVersion: 1, sessionToken: "fixture-session-token-123456", workbookConnectionId: "0123456789abcdef", serverTime: new Date().toISOString(), addinBuildId: "none" }) }); await new Promise(resolve => setTimeout(resolve, 30)); }
  return socket;
}

describe("task pane start-up in each Office application", () => {
  beforeEach(() => { localStorage.clear(); });
  afterEach(() => { uninstallWordFake(); delete g.Office; delete g.WebSocket; });

  it("Word: connects as Word, brands itself and hides the Excel-only testing tools", async () => {
    installWordFake(new FakeWordDoc());
    const socket = await boot("Word");
    const hello = socket.sent.find((message) => message.type === "hello");
    expect(hello).toMatchObject({ type: "hello", clientInfo: { host: "word" } });
    expect(document.title).toBe("Office Partner");
    expect(document.querySelector(".brand-text strong")?.textContent).toBe("Office Partner");
    expect(document.body.dataset.host).toBe("word");
    expect((document.getElementById("files-button") as HTMLElement).hidden).toBe(false); // connected files work from every app
    expect((document.getElementById("menu-testing") as HTMLElement).hidden).toBe(true);
  });

  it("Word: runs a tool request from the bridge in the open document and answers it", async () => {
    const doc = new FakeWordDoc();
    doc.setParagraphs([["Intro", "Heading 1"], "Hello world."]);
    installWordFake(doc);
    const socket = await boot("Word");
    const reply = (message: unknown) => socket.onmessage?.({ data: JSON.stringify(message) });
    reply({ type: "hello_ok", protocolVersion: 1, sessionToken: "token", workbookConnectionId: "x", serverTime: new Date().toISOString(), addinBuildId: "b" });
    await new Promise((resolve) => setTimeout(resolve, 10));
    reply({ type: "tool_request", callId: "c1", requestId: "r1", tool: "read_doc_section", args: { count: 5 }, timeoutMs: 5000 });
    reply({ type: "tool_request", callId: "c2", requestId: "r1", tool: "read_range", args: { address: "A1" }, timeoutMs: 5000 });
    await new Promise((resolve) => setTimeout(resolve, 50));
    const results = socket.sent.filter((message) => message.type === "tool_result");
    const first = results.find((message) => message.callId === "c1");
    expect(first?.result.ok).toBe(true);
    expect(JSON.stringify(first?.result.data)).toContain("Hello world.");
    expect(results.find((message) => message.callId === "c2")?.result).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
  });

  it("PowerPoint: connects as PowerPoint", async () => {
    const installed = installFakePowerPoint(new FakeDeck());
    try {
      const socket = await boot("PowerPoint");
      expect(socket.sent.find((message) => message.type === "hello")).toMatchObject({ clientInfo: { host: "powerpoint" } });
      expect(document.body.dataset.host).toBe("powerpoint");
    } finally { installed.restore(); }
  });

  it("returns from a settings child to Settings and keeps Settings open when switching provider", async () => {
    installWordFake(new FakeWordDoc());
    await boot("Word");
    const get = (id: string) => document.getElementById(id)!;
    get("menu-settings").click();
    get("menu-skills").click();
    expect(get("panel-skills").hidden).toBe(false);
    expect(document.body.classList.contains("settings-active")).toBe(true);
    (get("panel-skills").querySelector("[data-close]") as HTMLButtonElement).click();
    expect(get("panel-settings").hidden).toBe(false);
    const provider = get("provider-select") as HTMLSelectElement;
    provider.value = "codex";
    provider.dispatchEvent(new Event("change"));
    expect(get("panel-settings").hidden).toBe(false);
    (get("panel-settings").querySelector("[data-close]") as HTMLButtonElement).click();
    expect(document.body.classList.contains("settings-active")).toBe(false);
  });

  it("Excel: connects as Excel and keeps the Excel-only controls", async () => {
    const socket = await boot("Excel");
    expect(socket.sent.find((message) => message.type === "hello")).toMatchObject({ clientInfo: { host: "excel" } });
    expect(document.body.dataset.host).toBe("excel");
    expect((document.getElementById("files-button") as HTMLElement).hidden).toBe(false);
    expect((document.getElementById("menu-testing") as HTMLElement).hidden).toBe(false);
  });

  it("does nothing in an Office application it has no tools for", async () => {
    vi.resetModules();
    document.documentElement.innerHTML = html.replace(/<script[^>]*><\/script>/g, "");
    let ready: ((info: { host: string }) => void) | undefined;
    g.Office = { HostType: { Excel: "Excel" }, onReady: (callback: any) => { ready = callback; }, context: { document: { settings: { get: () => null } } } };
    g.WebSocket = FakeSocket;
    FakeSocket.instances = [];
    await import("../addin/taskpane/main.js");
    ready?.({ host: "Outlook" });
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(FakeSocket.instances).toHaveLength(0);
  });
});

describe("which chat the pane opens", () => {
  const OLD_CHAT = "Remember this earlier question about the budget";

  function seedLastChat(): void {
    const now = Date.now();
    localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([{
      id: "earlier-chat", title: OLD_CHAT, provider: "claude", identityVersion: 1, host: "word", startedAt: now - 60_000, updatedAt: now - 30_000,
      messages: [{ kind: "user", text: OLD_CHAT }, { kind: "assistant", text: "Here is the answer." }],
      changes: [],
    }]));
  }

  beforeEach(() => { localStorage.clear(); sessionStorage.clear(); });
  afterEach(() => { uninstallWordFake(); delete g.Office; delete g.WebSocket; });

  it("starts a new chat when the pane is opened again, and keeps the old one in the history", async () => {
    seedLastChat();
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    expect(document.getElementById("messages")?.textContent ?? "").not.toContain(OLD_CHAT);
    expect(document.querySelector("#messages .welcome")).not.toBeNull();
    expect(JSON.parse(localStorage.getItem("excelPartner.sessions.v1") ?? "[]")).toHaveLength(1);
  });

  it("lists only the chats of the app it is open in", async () => {
    const now = Date.now();
    const chat = (id: string, host: string | undefined, text: string) => ({
      id, title: text, provider: "claude", identityVersion: 1, ...(host ? { host } : {}), startedAt: now - 60_000, updatedAt: now - 30_000,
      messages: [{ kind: "user", text }, { kind: "assistant", text: "done" }], changes: [],
    });
    localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([
      chat("e", "excel", "Excel budget question"), chat("w", "word", "Word intro rewrite"), chat("p", "powerpoint", "Slide four layout"), chat("old", undefined, "Chat from before apps were recorded"),
    ]));
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    (document.getElementById("menu-history") ?? document.querySelector<HTMLElement>("[data-action=history]"))?.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    const shown = document.getElementById("history-list")?.textContent ?? "";
    expect(shown).toContain("Word intro rewrite");
    for (const other of ["Excel budget question", "Slide four layout", "Chat from before apps were recorded"]) expect(shown).not.toContain(other);
    expect(JSON.parse(localStorage.getItem("excelPartner.sessions.v1") ?? "[]")).toHaveLength(4); // the other apps' chats are untouched
  });

  it("H24 restores this pane's exact chat instead of a newer chat from another pane", async () => {
    seedLastChat(); const source = JSON.parse(localStorage.getItem("excelPartner.sessions.v1")!);
    localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([...source, { ...source[0], id: "other-window", title: "Other window", updatedAt: Date.now(), messages: [{ kind: "user", text: "OTHER_WINDOW_NEWEST" }] }]));
    sessionStorage.setItem("excelPartner.restoreChat", "1"); sessionStorage.setItem("excelPartner.activeChatId", "earlier-chat");
    installWordFake(new FakeWordDoc()); await boot("Word", true);
    expect(document.getElementById("messages")?.textContent).toContain(OLD_CHAT);
    expect(document.getElementById("messages")?.textContent).not.toContain("OTHER_WINDOW_NEWEST");
  });

  it("brings the chat back when the pane reloaded itself, and only that once", async () => {
    seedLastChat();
    sessionStorage.setItem("excelPartner.restoreChat", "1");
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    expect(document.getElementById("messages")?.textContent ?? "").toContain(OLD_CHAT);
    expect(sessionStorage.getItem("excelPartner.restoreChat")).toBeNull();
    uninstallWordFake();
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    expect(document.getElementById("messages")?.textContent ?? "").not.toContain(OLD_CHAT);
  });

  it("continues a saved chat and deletes the current chat through History", async () => {
    seedLastChat();
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    document.getElementById("history-button")!.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    document.querySelector<HTMLButtonElement>(".history-item")!.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    expect(document.getElementById("session-transcript")!.textContent).toContain(OLD_CHAT);
    document.getElementById("session-continue")!.click();
    await vi.waitFor(() => expect(document.getElementById("messages")!.textContent).toContain(OLD_CHAT));
    expect(localStorage.getItem("excelPartner.provider")).toBe("claude");
    document.getElementById("history-button")!.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    expect(document.querySelector<HTMLButtonElement>(".history-item")!.dataset.current).toBe("true");
    document.querySelector<HTMLButtonElement>(".history-delete")!.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    expect(document.querySelector("#messages .welcome")).not.toBeNull();
    expect(JSON.parse(localStorage.getItem("excelPartner.sessions.v1") ?? "[]")).toHaveLength(0);
  });

  it("requires the second click to clear history and resets the current chat", async () => {
    seedLastChat();
    sessionStorage.setItem("excelPartner.restoreChat", "1");
    installWordFake(new FakeWordDoc());
    await boot("Word", true);
    const clear = document.getElementById("history-clear")!;
    clear.click();
    expect(document.getElementById("messages")!.textContent).toContain(OLD_CHAT);
    expect(JSON.parse(localStorage.getItem("excelPartner.sessions.v1") ?? "[]")).toHaveLength(1);
    clear.click();
    await new Promise(resolve => setTimeout(resolve, 30));
    expect(document.querySelector("#messages .welcome")).not.toBeNull();
    expect(JSON.parse(localStorage.getItem("excelPartner.sessions.v1") ?? "[]")).toHaveLength(0);
    expect(clear.dataset.confirm).toBe("0");
  });
});

describe("model picker across provider changes", () => {
  beforeEach(() => { localStorage.clear(); sessionStorage.clear(); });
  afterEach(() => { uninstallWordFake(); delete g.Office; delete g.WebSocket; });

  it("keeps a late catalog response with its provider and restores that provider's saved model", async () => {
    localStorage.setItem("excelPartner.models.v1", JSON.stringify({ copilot: "copilot-model", codex: "codex-model" }));
    installWordFake(new FakeWordDoc());
    const socket = await boot("Word");
    let resolveCopilot!: (response: Response) => void;
    g.fetch = vi.fn(async (input: string) => {
      if (input.includes("/session/models?provider=copilot")) return new Promise<Response>(resolve => { resolveCopilot = resolve; });
      if (input.includes("/session/models?provider=codex")) return Response.json({ models: [{ id: "codex-model", name: "Codex model" }] });
      return Response.json({ sessions: [], commands: [], skills: [], servers: [], enabled: [], roots: [] });
    });
    socket.onmessage?.({ data: JSON.stringify({ type: "hello_ok", protocolVersion: 1, sessionToken: "test-session-token", workbookConnectionId: "x", serverTime: new Date().toISOString(), addinBuildId: "b" }) });
    const provider = document.getElementById("provider-select") as HTMLSelectElement;
    provider.value = "codex";
    provider.dispatchEvent(new Event("change"));
    const model = document.getElementById("model-select") as HTMLSelectElement;
    await vi.waitFor(() => expect(model.value).toBe("codex-model"));
    resolveCopilot(Response.json({ models: [{ id: "copilot-model", name: "Copilot model" }] }));
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(model.value).toBe("codex-model");
    provider.value = "copilot";
    provider.dispatchEvent(new Event("change"));
    expect(model.value).toBe("copilot-model");
    expect(document.getElementById("model-button")!.textContent).toContain("Copilot model");
  });
});
