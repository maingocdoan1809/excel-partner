import { describe, expect, it, vi } from "vitest";
import { ProviderLoginManager } from "../bridge/src/provider-login.js";
import type { ProviderLoginStatus } from "../protocol/src/provider-login.js";
import type { ProviderId } from "../protocol/src/index.js";
import type { ProviderRegistry } from "../providers/src/registry.js";

/** A sign-in attempt the test controls: the manager only sees the interface every provider implements. */
function fakeAttempt(provider: ProviderId) {
  let current: ProviderLoginStatus = { state: "starting", provider, userCode: null, verificationUri: null, expiresAt: null, message: "Getting a sign-in link…" };
  const attempt = {
    status: () => ({ ...current }), submitCode: vi.fn(),
    cancel: vi.fn(() => attempt.set({ state: "cancelled", message: "Sign-in cancelled." })),
    set(next: Partial<ProviderLoginStatus>) { current = { ...current, ...next }; },
  };
  return attempt;
}

function setup() {
  const attempts: ReturnType<typeof fakeAttempt>[] = [];
  const kinds: Partial<Record<ProviderId, "device" | "account">> = { opencode: "account" };
  const adapters = new Map<ProviderId, { capabilities: { loginKind: string }; startLogin: ReturnType<typeof vi.fn>; logout: ReturnType<typeof vi.fn>; loginAccounts?: ReturnType<typeof vi.fn> }>();
  const lookup = (id: ProviderId) => {
    if (!adapters.has(id)) {
      adapters.set(id, {
        capabilities: { loginKind: kinds[id] ?? "device" },
        startLogin: vi.fn(() => { const attempt = fakeAttempt(id); attempts.push(attempt); return attempt; }),
        logout: vi.fn(async () => {}),
        ...(id === "opencode" ? { loginAccounts: vi.fn(async () => [{ id: "openai", name: "OpenAI", connected: false, methods: [] }]) } : {}),
      });
    }
    return adapters.get(id)!;
  };
  const registry = { get: lookup, require: lookup } as unknown as ProviderRegistry;
  return { manager: new ProviderLoginManager(registry), attempts, adapters };
}

describe("the bridge's sign-in manager", () => {
  it("starts the provider's own sign-in and reports its status", () => {
    const { manager, attempts } = setup();
    expect(manager.start("window", "codex").state).toBe("starting");
    attempts[0]!.set({ state: "waiting", userCode: "AB12-CD34", verificationUri: "https://example.com/device" });
    expect(manager.status("window", "codex")).toMatchObject({ state: "waiting", userCode: "AB12-CD34" });
  });

  it("lets only one window sign in at a time, and answers a repeated start with the running one", () => {
    const { manager, adapters } = setup();
    manager.start("first", "copilot");
    manager.start("first", "copilot");
    expect(adapters.get("copilot")!.startLogin).toHaveBeenCalledTimes(1);
    expect(() => manager.start("second", "copilot")).toThrow("Another window");
    expect(manager.status("second", "copilot").state).toBe("idle");
    manager.cancel("second");
    expect(manager.status("first", "copilot").state).toBe("starting");
    expect(manager.isBusy).toBe(true);
  });

  it("cancels the earlier attempt of the same window when it switches provider or picks an account", () => {
    const { manager, attempts } = setup();
    manager.start("window", "copilot");
    expect(manager.start("window", "codex").provider).toBe("codex");
    expect(attempts[0]!.cancel).toHaveBeenCalledOnce();
    manager.start("window", "opencode", { accountId: "openai", method: 0 });
    manager.start("window", "opencode", { accountId: "poe", method: 0 });
    expect(attempts[2]!.cancel).toHaveBeenCalledOnce();
  });

  it("does not replay a finished sign-in when the screen is opened again, but keeps one still running", () => {
    const { manager, attempts } = setup();
    manager.start("window", "codex");
    attempts[0]!.set({ state: "success", message: "Signed in." });
    expect(manager.status("window", "codex").state).toBe("success");
    expect(manager.status("window", "codex", true)).toMatchObject({ state: "idle", provider: "codex" });
    expect(manager.status("window", "codex").state).toBe("idle");

    manager.start("window", "codex");
    attempts[1]!.set({ state: "waiting", userCode: "ABCD-EFGH" });
    expect(manager.status("window", "codex", true)).toMatchObject({ state: "waiting", userCode: "ABCD-EFGH" });
  });

  it("says what to do first according to how the provider signs in", () => {
    const { manager } = setup();
    expect(manager.status("window", "codex").message).toBe("Connect your account to get started.");
    expect(manager.status("window", "opencode").message).toBe("Choose the account you want to use.");
  });

  it("hands the code to the attempt, and refuses one when nothing is waiting for it", () => {
    const { manager, attempts } = setup();
    expect(() => manager.submitCode("window", "abcd1234")).toThrow("no sign-in waiting");
    manager.start("window", "claude");
    attempts[0]!.set({ state: "waiting" });
    manager.submitCode("window", "abcd1234");
    expect(attempts[0]!.submitCode).toHaveBeenCalledWith("abcd1234");
  });

  it("lists the accounts of a provider that has them, and none for one that does not", async () => {
    const { manager } = setup();
    expect(await manager.accounts("opencode")).toMatchObject([{ id: "openai" }]);
    expect(await manager.accounts("codex")).toEqual([]);
  });

  it("signs a provider out, but not while someone is signing in", async () => {
    const { manager, adapters } = setup();
    await manager.logout("codex");
    expect(adapters.get("codex")!.logout).toHaveBeenCalledWith(undefined);
    await manager.logout("opencode", "openai");
    expect(adapters.get("opencode")!.logout).toHaveBeenCalledWith("openai");
    manager.start("window", "copilot");
    await expect(manager.logout("codex")).rejects.toThrow("Finish or cancel");
  });

  it("cancels what a closed window was signing in", () => {
    const { manager, attempts } = setup();
    manager.start("window", "codex");
    manager.forget("window");
    expect(attempts[0]!.cancel).toHaveBeenCalledOnce();
    expect(manager.status("window", "codex").state).toBe("idle");
    manager.dispose();
  });
});
