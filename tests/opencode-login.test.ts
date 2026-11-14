import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { OpenCodeAuthError, OpenCodeLoginAttempt } from "../providers/src/opencode/auth.js";
import { OpencodeAdapter } from "../providers/src/opencode/adapter.js";

function fakeService() {
  return {
    closed: false,
    catalog: vi.fn(async () => []),
    authorize: vi.fn(async () => ({ url: "https://auth.openai.com/codex/device", method: "auto" as const, instructions: "Enter code: G5PX-R5UGP" })),
    complete: vi.fn(async (_account: string, _method: number, _code?: string) => {}),
    setKey: vi.fn(async (_account: string, _key: string, _metadata?: Record<string, string>) => {}),
    close: vi.fn(),
  };
}

describe("signing in to an OpenCode account through OpenCode's own service", () => {
  const flush = () => vi.advanceTimersByTimeAsync(0);
  let service: ReturnType<typeof fakeService>;
  beforeEach(() => { vi.useFakeTimers(); service = fakeService(); });
  afterEach(() => vi.useRealTimers());

  it("shows the page and code, then succeeds once OpenCode has saved the account", async () => {
    let finish!: () => void;
    service.complete.mockImplementation(() => new Promise<void>(resolve => { finish = resolve; }));
    const attempt = OpenCodeLoginAttempt.start(service, { accountId: "openai", method: 1 });
    expect(attempt.status().state).toBe("starting");
    await flush();
    expect(service.authorize).toHaveBeenCalledWith("openai", 1, {});
    expect(attempt.status()).toMatchObject({ state: "waiting", provider: "opencode", userCode: "G5PX-R5UGP", verificationUri: "https://auth.openai.com/codex/device", needsCode: false });
    finish();
    await flush();
    expect(attempt.status()).toMatchObject({ state: "success", userCode: null });
    expect(service.close).toHaveBeenCalled();
  });

  it("passes the answers to the provider's questions along and asks for the pasted code when the page returns one", async () => {
    service.authorize.mockResolvedValue({ url: "https://github.com/login/oauth", method: "code" as never, instructions: "Paste the code" });
    const attempt = OpenCodeLoginAttempt.start(service, { accountId: "github-copilot", method: 0, inputs: { deploymentType: "enterprise", enterpriseUrl: "company.ghe.com" } });
    await flush();
    expect(service.authorize).toHaveBeenCalledWith("github-copilot", 0, { deploymentType: "enterprise", enterpriseUrl: "company.ghe.com" });
    expect(attempt.status()).toMatchObject({ state: "waiting", needsCode: true, userCode: null });
    expect(() => attempt.submitCode("bad code!")).toThrow("sign-in code");
    attempt.submitCode(" abc123#xyz ");
    expect(attempt.status()).toMatchObject({ state: "waiting", needsCode: false });
    expect(service.complete).toHaveBeenCalledWith("github-copilot", 0, "abc123#xyz");
    await flush();
    expect(attempt.status().state).toBe("success");
  });

  it("reports a code the provider does not accept", async () => {
    service.authorize.mockResolvedValue({ url: "https://example.com/login", method: "code" as never, instructions: "" });
    service.complete.mockRejectedValue(new Error("nope"));
    const attempt = OpenCodeLoginAttempt.start(service, { accountId: "poe", method: 0 });
    await flush();
    attempt.submitCode("abc123");
    await flush();
    expect(attempt.status()).toMatchObject({ state: "error", message: expect.stringContaining("not accepted") });
  });

  it("hands an API key straight to OpenCode without keeping or returning it", async () => {
    const attempt = OpenCodeLoginAttempt.start(service, { accountId: "anthropic", method: 0, key: "  sk-secret-value  " });
    const first = attempt.status();
    await flush();
    expect(service.setKey).toHaveBeenCalledWith("anthropic", "sk-secret-value", undefined);
    expect(service.authorize).not.toHaveBeenCalled();
    expect(JSON.stringify([first, attempt.status()])).not.toContain("sk-secret-value");
    expect(attempt.status().state).toBe("success");
  });

  it("sends the extra answers some key providers need as metadata", async () => {
    OpenCodeLoginAttempt.start(service, { accountId: "cloudflare-workers-ai", method: 0, key: "token", inputs: { accountId: "abc123" } });
    await flush();
    expect(service.setKey).toHaveBeenCalledWith("cloudflare-workers-ai", "token", { accountId: "abc123" });
  });

  it("rejects an unsafe account id, a bad method or a key with control characters before starting anything", () => {
    expect(() => OpenCodeLoginAttempt.start(service, { accountId: "../x", method: 0 })).toThrow("Choose an account");
    expect(() => OpenCodeLoginAttempt.start(service, { accountId: "openai", method: -1 })).toThrow("Choose an account");
    expect(() => OpenCodeLoginAttempt.start(service, { accountId: "openai", method: 0, key: "a\nb" })).toThrow("API key");
    expect(service.authorize).not.toHaveBeenCalled();
  });

  it("closes the service on cancel and shows OpenCode's own message when it refuses", async () => {
    service.complete.mockImplementation(() => new Promise<void>(() => {}));
    const attempt = OpenCodeLoginAttempt.start(service, { accountId: "openai", method: 1 });
    await flush();
    attempt.cancel();
    expect(attempt.status().state).toBe("cancelled");
    expect(service.close).toHaveBeenCalled();

    service.authorize.mockRejectedValue(new OpenCodeAuthError("Enter a valid GitHub Enterprise URL."));
    const refused = OpenCodeLoginAttempt.start(service, { accountId: "github-copilot", method: 0 });
    await flush();
    expect(refused.status()).toMatchObject({ state: "error", message: "Enter a valid GitHub Enterprise URL." });
  });

  it("is what the OpenCode adapter starts for a chosen account, and needs a choice", () => {
    const adapter = new OpencodeAdapter({ env: {}, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    vi.spyOn(adapter as unknown as { openAuthService(): unknown }, "openAuthService").mockReturnValue(service);
    expect(() => adapter.startLogin()).toThrow("Choose an account");
    expect(adapter.startLogin({ accountId: "openai", method: 1 }).status().provider).toBe("opencode");
  });
});
