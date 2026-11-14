import { describe, expect, it, vi } from "vitest";
import type { ProviderDetection } from "../protocol/src/index.js";
import { BaseCliAdapter, type CliProbe } from "../providers/src/base.js";
import { ProviderLoginError } from "../providers/src/login.js";
import { MockAdapter } from "../providers/src/mock.js";
import type { McpEndpoint, StartSessionContext } from "../providers/src/types.js";

const mcp: McpEndpoint = { bridgeUrl: "http://127.0.0.1:1", sessionId: "preflight", token: "t", command: "node", args: [], env: {} };

class Fake extends BaseCliAdapter {
  readonly id = "codex" as const;
  readonly capabilities = MockAdapter.capabilities;
  readonly started: StartSessionContext[] = [];
  readonly closed: string[] = [];
  constructor(private readonly detection: Partial<ProviderDetection>) { super({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } }); }
  protected probe(): CliProbe { return { command: "fake", versionArgs: [], installHint: "", loginHint: "" }; }
  override async detect(): Promise<ProviderDetection> { return { provider: "codex", installed: true, version: "1", status: "ready", hint: null, capabilities: this.capabilities, ...this.detection }; }
  override async startSession(ctx: StartSessionContext) { this.started.push(ctx); return "native"; }
  override async closeSession(id: string) { this.closed.push(id); }
  override async *sendTurn(): AsyncIterable<never> { /* no turns here */ }
  override async cancelTurn() {}
}

describe("proving a provider is ready for a turn", () => {
  it("opens and closes a session with the given tools once the CLI is signed in, without sending a prompt", async () => {
    const adapter = new Fake({});
    await adapter.verifyReady(mcp);
    expect(adapter.started).toHaveLength(1);
    expect(adapter.started[0]).toMatchObject({ workbookConnectionId: "preflight", mcp, model: null });
    expect(adapter.closed).toEqual(["native"]);
  });

  it("asks the person to sign in, without opening a session, when the CLI is signed out", async () => {
    const adapter = new Fake({ status: "login_required" });
    await expect(adapter.verifyReady(mcp)).rejects.toBeInstanceOf(ProviderLoginError);
    expect(adapter.started).toHaveLength(0);
  });

  it("does not open a session for a CLI that is missing or broken", async () => {
    const adapter = new Fake({ installed: false, status: "not_installed" });
    await expect(adapter.verifyReady(mcp)).rejects.toThrow("provider initialization failed");
    expect(vi.mocked(adapter.started)).toHaveLength(0);
  });
});
