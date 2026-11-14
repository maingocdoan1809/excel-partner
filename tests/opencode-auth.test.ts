import { EventEmitter } from "node:events";
import { describe, expect, it, vi } from "vitest";
import { OpenCodeAuthServer } from "../providers/src/opencode/auth.js";

function fakeChild() {
  return Object.assign(new EventEmitter(), {
    stdout: Object.assign(new EventEmitter(), { resume: vi.fn() }), stderr: Object.assign(new EventEmitter(), { resume: vi.fn() }),
    pid: undefined, exitCode: null as number | null, signalCode: null, kill: vi.fn(),
  });
}

const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

function setup(routes: Record<string, (init: RequestInit) => Response>) {
  const child = fakeChild();
  const spawn = vi.fn(() => child);
  const calls: { method: string; path: string; body: unknown; headers: Record<string, string> }[] = [];
  const fetchMock = vi.fn(async (input: string | URL | Request, init: RequestInit = {}) => {
    const path = new URL(String(input)).pathname;
    calls.push({ method: init.method ?? "GET", path, body: init.body ? JSON.parse(String(init.body)) : undefined, headers: init.headers as Record<string, string> });
    const route = routes[`${init.method ?? "GET"} ${path}`];
    return route ? route(init) : json({}, 404);
  });
  const server = new OpenCodeAuthServer({ command: "opencode-test", env: { KEEP: "1" } }, { spawn: spawn as never, fetch: fetchMock as never, port: async () => 4321 });
  return { server, child, spawn, calls };
}

const authList = {
  openai: [{ type: "oauth", label: "ChatGPT Pro/Plus (browser)" }, { type: "oauth", label: "ChatGPT Pro/Plus (headless)" }, { type: "api", label: "Manually enter API Key" }],
  "github-copilot": [{
    type: "oauth", label: "Login with GitHub Copilot", prompts: [
      { type: "select", key: "deploymentType", message: "Select GitHub deployment type", options: [{ label: "GitHub.com", value: "github.com", hint: "Public" }, { label: "GitHub Enterprise", value: "enterprise" }] },
      { type: "text", key: "enterpriseUrl", message: "Enter your GitHub Enterprise URL or domain", placeholder: "company.ghe.com", when: { key: "deploymentType", op: "eq", value: "enterprise" } },
    ],
  }],
};

describe("OpenCode's sign-in service", () => {
  it("starts a private loopback server with a random password, never on a shared port", async () => {
    const { server, spawn, calls } = setup({ "GET /provider/auth": () => json(authList), "GET /provider": () => json({ all: [], connected: [] }) });
    await server.catalog();
    expect(spawn).toHaveBeenCalledWith("opencode-test", ["serve", "--port", "4321", "--hostname", "127.0.0.1"],
      expect.objectContaining({ windowsHide: true, env: expect.objectContaining({ KEEP: "1", OPENCODE_SERVER_USERNAME: "office-partner", OPENCODE_SERVER_PASSWORD: expect.stringMatching(/^[0-9a-f]{36}$/) }) }));
    expect(calls[0]!.headers.authorization).toMatch(/^Basic /);
    server.close();
  });

  it("lists every account, with the methods OpenCode offers and a plain API key for the rest", async () => {
    const { server } = setup({
      "GET /provider/auth": () => json(authList),
      "GET /provider": () => json({ all: [{ id: "anthropic", name: "Anthropic" }, { id: "openai", name: "OpenAI" }, { id: "github-copilot", name: "GitHub Copilot" }], connected: ["openai"] }),
    });
    const accounts = await server.catalog();
    expect(accounts.map(account => account.id)).toEqual(["openai", "github-copilot", "anthropic"]);
    expect(accounts.find(account => account.id === "openai")).toMatchObject({ connected: true, methods: [{ index: 0, type: "oauth" }, { index: 1, type: "oauth" }, { index: 2, type: "api" }] });
    expect(accounts.find(account => account.id === "anthropic")).toMatchObject({ connected: false, methods: [{ index: 0, type: "api", label: "API key", prompts: [] }] });
    const copilot = accounts.find(account => account.id === "github-copilot")!;
    expect(copilot.methods[0]!.prompts).toEqual([
      { key: "deploymentType", type: "select", message: "Select GitHub deployment type", options: [{ label: "GitHub.com", value: "github.com", hint: "Public" }, { label: "GitHub Enterprise", value: "enterprise" }] },
      { key: "enterpriseUrl", type: "text", message: "Enter your GitHub Enterprise URL or domain", placeholder: "company.ghe.com", when: { key: "deploymentType", op: "eq", value: "enterprise" } },
    ]);
    server.close();
  });

  it("authorizes, completes with a pasted code and saves an API key through OpenCode's own routes", async () => {
    const { server, calls } = setup({
      "GET /provider/auth": () => json({}),
      "POST /provider/openai/oauth/authorize": () => json({ url: "https://auth.openai.com/codex/device", method: "auto", instructions: "Enter code: AB12-CD34" }),
      "POST /provider/openai/oauth/callback": () => json(true),
      "PUT /auth/anthropic": () => json(true),
    });
    expect(await server.authorize("openai", 1, { a: "b" })).toEqual({ url: "https://auth.openai.com/codex/device", method: "auto", instructions: "Enter code: AB12-CD34" });
    await server.complete("openai", 1, "pasted");
    await server.setKey("anthropic", "sk-1", { accountId: "x" });
    expect(calls.filter(call => call.method !== "GET").map(call => [call.method, call.path, call.body])).toEqual([
      ["POST", "/provider/openai/oauth/authorize", { method: 1, inputs: { a: "b" } }],
      ["POST", "/provider/openai/oauth/callback", { method: 1, code: "pasted" }],
      ["PUT", "/auth/anthropic", { type: "api", key: "sk-1", metadata: { accountId: "x" } }],
    ]);
    server.close();
  });

  it("surfaces OpenCode's own explanation when it refuses", async () => {
    const { server } = setup({
      "GET /provider/auth": () => json({}),
      "POST /provider/github-copilot/oauth/authorize": () => json({ name: "ProviderAuthValidationFailed", data: { message: "Enter a valid URL." } }, 400),
    });
    await expect(server.authorize("github-copilot", 0, {})).rejects.toThrow("Enter a valid URL.");
    server.close();
  });

  it("stops the server when closed, and reports a CLI that cannot run", async () => {
    const { server, child } = setup({ "GET /provider/auth": () => json({}) });
    await server.authorize("x", 0, {}).catch(() => {});
    server.close();
    expect(child.kill).toHaveBeenCalled();

    const broken = setup({});
    const attempt = broken.server.catalog();
    attempt.catch(() => {});
    await vi.waitFor(() => expect(broken.spawn).toHaveBeenCalled());
    broken.child.emit("error", new Error("ENOENT"));
    await expect(attempt).rejects.toThrow("Could not run the OpenCode CLI");
  });
});
