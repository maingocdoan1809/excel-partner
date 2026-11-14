import { DEFAULT_PROVIDER, type ProviderId } from "../../protocol/src/providers.js";
import type { LoginAccount, LoginChoice, ProviderLoginStatus } from "../../protocol/src/provider-login.js";
import { idleLoginStatus, ProviderLoginError } from "../../providers/src/login.js";
import type { ProviderRegistry } from "../../providers/src/registry.js";
import type { ProviderLoginAttempt } from "../../providers/src/types.js";

export { ProviderLoginError };

interface Entry { provider: ProviderId; attempt: ProviderLoginAttempt }

/** Keeps one sign-in per pane window and lets only one window sign in at a time. The provider does the signing in. */
export class ProviderLoginManager {
  private readonly entries = new Map<string, Entry>();

  constructor(private readonly providers: ProviderRegistry) {}

  get isBusy(): boolean { return [...this.entries.values()].some(entry => this.isActive(entry)); }

  status(sessionId: string, provider: ProviderId = DEFAULT_PROVIDER, fresh = false): ProviderLoginStatus {
    const entry = this.entries.get(sessionId);
    if (fresh && entry && !this.isActive(entry)) this.entries.delete(sessionId);
    else if (entry?.provider === provider) return entry.attempt.status();
    return this.idle(provider);
  }

  private idle(provider: ProviderId): ProviderLoginStatus {
    return idleLoginStatus(provider, this.providers.get(provider)?.capabilities.loginKind === "account"
      ? "Choose the account you want to use." : "Connect your account to get started.");
  }

  /** The accounts to choose from, for providers that list them. */
  accounts(provider: ProviderId): Promise<LoginAccount[]> {
    return this.providers.require(provider).loginAccounts?.() ?? Promise.resolve([]);
  }

  /** Signs the provider out, unless a sign-in is going on. */
  async logout(provider: ProviderId, accountId?: string): Promise<void> {
    if (this.isBusy) throw new ProviderLoginError("Finish or cancel the sign-in first.");
    const adapter = this.providers.get(provider);
    if (!adapter) throw new ProviderLoginError("That provider is not available.");
    await adapter.logout(accountId);
  }

  start(sessionId: string, provider: ProviderId = DEFAULT_PROVIDER, choice?: LoginChoice): ProviderLoginStatus {
    for (const [owner, entry] of this.entries) {
      if (!this.isActive(entry)) continue;
      if (owner !== sessionId) throw new ProviderLoginError("Another window is signing in. Finish or cancel that session first.");
      if (entry.provider === provider && !choice) return entry.attempt.status();
      entry.attempt.cancel();
    }
    const entry = { provider, attempt: this.providers.require(provider).startLogin(choice) };
    this.entries.set(sessionId, entry);
    return entry.attempt.status();
  }

  /** Hands the code the sign-in page showed to the attempt. */
  submitCode(sessionId: string, code: string): ProviderLoginStatus {
    const entry = this.entries.get(sessionId);
    if (!entry || !this.isActive(entry)) throw new ProviderLoginError("There is no sign-in waiting for a code.");
    entry.attempt.submitCode(code);
    return entry.attempt.status();
  }

  cancel(sessionId: string): ProviderLoginStatus {
    const entry = this.entries.get(sessionId);
    if (!entry) return this.idle(DEFAULT_PROVIDER);
    if (this.isActive(entry)) entry.attempt.cancel();
    return entry.attempt.status();
  }

  forget(sessionId: string): void {
    this.cancel(sessionId);
    this.entries.delete(sessionId);
  }

  dispose(): void {
    for (const sessionId of [...this.entries.keys()]) this.forget(sessionId);
  }

  private isActive(entry: Entry): boolean {
    const { state } = entry.attempt.status();
    return state === "starting" || state === "waiting";
  }
}
