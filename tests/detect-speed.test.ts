import { describe, expect, it } from "vitest";
import { BaseCliAdapter, type CliProbe } from "../providers/src/base.js";
import { MockAdapter } from "../providers/src/mock.js";
import type { ProviderConnectionStatus } from "../protocol/src/index.js";

const DELAY_MS = 400;

/** A provider whose two checks each take a while, as a large CLI does: it takes about a second to start. */
class SlowCheckAdapter extends BaseCliAdapter {
  readonly id = "claude" as const;
  readonly capabilities = MockAdapter.capabilities;
  loginChecks = 0;
  constructor(private readonly versionCode: number, private readonly login: ProviderConnectionStatus = "ready") {
    super({ bridgeUrl: "http://127.0.0.1:1", mcp: { command: process.execPath, args: [] }, logger: { debug() {}, warn() {} } });
  }
  protected override probe(): CliProbe {
    return { command: process.execPath, versionArgs: ["-e", `setTimeout(() => { console.log("9.9.9"); process.exit(${this.versionCode}); }, ${DELAY_MS})`], installHint: "install it", loginHint: "login" };
  }
  protected override async checkLogin() {
    this.loginChecks += 1;
    await new Promise(resolve => setTimeout(resolve, DELAY_MS));
    return { status: this.login, hint: this.login === "ready" ? null : "sign in" };
  }
  override async startSession(): Promise<string> { throw new Error("not used"); }
  // eslint-disable-next-line require-yield
  override async *sendTurn(): AsyncIterable<never> { throw new Error("not used"); }
  override async cancelTurn(): Promise<void> {}
  override async closeSession(): Promise<void> {}
}

describe("checking a provider CLI", () => {
  it("asks for the version and the login at the same time", async () => {
    const adapter = new SlowCheckAdapter(0);
    const started = performance.now();
    const detection = await adapter.detect(true);
    const elapsed = performance.now() - started;
    expect(detection).toMatchObject({ installed: true, status: "ready" });
    // Two checks of 400 ms one after the other would take at least 800 ms.
    expect(elapsed).toBeLessThan(DELAY_MS * 1.8);
  });

  it("still reports the login that was found", async () => {
    const detection = await new SlowCheckAdapter(0, "login_required").detect(true);
    expect(detection).toMatchObject({ installed: true, status: "login_required", hint: "sign in" });
  });

  it("ignores the login of a CLI that did not run", async () => {
    const adapter = new SlowCheckAdapter(1, "ready");
    const detection = await adapter.detect(true);
    expect(detection.status).toBe("error");
    expect(detection.hint).not.toBe(null);
  });

  it("keeps the answer for a moment instead of running the CLI again", async () => {
    const adapter = new SlowCheckAdapter(0);
    await adapter.detect(true);
    const started = performance.now();
    await adapter.detect();
    expect(performance.now() - started).toBeLessThan(50);
    expect(adapter.loginChecks).toBe(1);
  });
});
