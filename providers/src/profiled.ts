import path from "node:path";
import fs from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import type { ProviderCapabilities, ProviderEvent, ProviderId } from "../../protocol/src/index.js";
import type { LoginAccount, LoginChoice } from "../../protocol/src/provider-login.js";
import type { McpEndpoint, ProviderAdapter, ProviderLoginAttempt, ProviderRuntime, SendTurnInput, StartSessionContext } from "./types.js";

export function providerPackageRoot(): string {
  let root = path.dirname(fileURLToPath(import.meta.url));
  for (let i = 0; i < 8; i++, root = path.dirname(root)) if (fs.existsSync(path.join(root, "package.json"))) return root;
  throw new Error("Provider package root not found");
}
const store = createRequire(import.meta.url)(path.join(providerPackageRoot(), "bin", "runtime-store.cjs"));

/** Oldest CLI version the release has been verified against, from the one manifest that also pins the installed version. */
export function minimumVersion(provider: ProviderId): string | null {
  return store.manifest.providers[provider]?.minVersion ?? null;
}

/** Installs the CLI the release names for the provider, checks it and makes the provider use it. Slow: it downloads. */
export async function installProvider(provider: ProviderId, env: NodeJS.ProcessEnv = process.env): Promise<void> {
  const dataDir = store.home(env), root = providerPackageRoot();
  const tools = (name: string) => import(pathToFileURL(path.join(root, "bin", name)).href);
  (await tools("manage.mjs")).ensureBaseline(root, dataDir);
  await (await tools("runtime-manager.mjs")).installRuntime(root, provider, { dataDir });
}

/** Stops the provider from using the CLI the product installed for it. */
export async function uninstallProvider(provider: ProviderId, env: NodeJS.ProcessEnv = process.env): Promise<void> {
  (await import(pathToFileURL(path.join(providerPackageRoot(), "bin", "runtime-manager.mjs")).href)).uninstallRuntime(provider, { dataDir: store.home(env) });
}

/** An adapter class together with what it can do, known before any CLI is looked at. */
export interface AdapterClass { new (runtime: ProviderRuntime): ProviderAdapter; readonly capabilities: ProviderCapabilities; readonly clientToolNotes?: string }

/**
 * Runs a provider's adapter against the CLI the product installed for it, or the system CLI when it installed none.
 * The install is looked up on every call, so a CLI installed or switched while the bridge runs is used by the next one.
 */
export class ProfiledAdapter implements ProviderAdapter {
  private readonly adapters = new Map<string, ProviderAdapter>();
  /** The adapter that opened each session, which stays the one that answers for it. */
  private readonly sessions = new Map<string, ProviderAdapter>();
  private readonly dataDir: string;
  constructor(readonly id: ProviderAdapter["id"], private readonly runtime: ProviderRuntime, private readonly adapterClass: AdapterClass) {
    this.dataDir = store.home(runtime.env ?? process.env);
  }
  get capabilities() { return this.adapterClass.capabilities; }
  get clientToolNotes() { return this.adapterClass.clientToolNotes; }
  runtimeRevision() { const state = store.readState(this.dataDir); return state.disabled.includes(this.id) ? "disabled" : state.active[this.id] ?? "system"; }

  private resolve(): ProviderAdapter {
    const profile = store.selectProfile(this.dataDir, this.id);
    const key = profile?.id ?? "system";
    let adapter = this.adapters.get(key);
    if (!adapter) {
      adapter = new this.adapterClass({ ...this.runtime, env: store.childEnvironment(this.runtime.env ?? process.env, profile) });
      this.adapters.set(key, adapter);
    }
    return adapter;
  }
  async detect(force = false) {
    try { return await this.resolve().detect(force); }
    catch (error) {
      const disabled = error instanceof Error && error.message === "PROVIDER_DISABLED";
      return { provider: this.id, installed: !disabled, version: null, status: disabled ? "not_installed" as const : "error" as const,
        hint: disabled ? `Install this provider with office-partner providers install ${this.id}.` : "The installed provider runtime is unavailable. Run office-partner doctor and office-partner repair.",
        capabilities: this.capabilities };
    }
  }
  startLogin(choice?: LoginChoice): ProviderLoginAttempt { return this.resolve().startLogin(choice); }
  async loginAccounts(): Promise<LoginAccount[]> { return (await this.resolve().loginAccounts?.()) ?? []; }
  async logout(accountId?: string) { await this.resolve().logout(accountId); }
  /** Installing does not look at the current install, so it works for a provider that has none or is turned off. */
  install() { return installProvider(this.id, this.runtime.env ?? process.env); }
  uninstall() { return uninstallProvider(this.id, this.runtime.env ?? process.env); }
  async verifyReady(mcp: McpEndpoint) { return this.resolve().verifyReady(mcp); }
  async connectionStatus() { return this.resolve().connectionStatus(); }
  async listModels() { return (await this.resolve().listModels?.()) ?? { models: [], defaultModelId: null }; }
  async listCommands(id?: string) { return ((id && this.sessions.get(id)) || this.resolve()).listCommands?.(id) ?? []; }
  async sessionTitle(nativeId: string) { return this.resolve().sessionTitle?.(nativeId); }
  async startSession(ctx: StartSessionContext) {
    const adapter = this.resolve();
    let starting = true, closed = false, accepted = false;
    const id = await adapter.startSession({ ...ctx, onRuntimeClosed: () => {
      if (starting) closed = true;
      else if (accepted) ctx.onRuntimeClosed?.();
    } });
    starting = false;
    if (closed) { await adapter.closeSession(id); throw new Error("Provider runtime closed during initialization"); }
    accepted = true;
    this.sessions.set(id, adapter);
    return id;
  }
  private worker(id: string): ProviderAdapter {
    const adapter = this.sessions.get(id);
    if (!adapter) throw new Error("Provider worker is not available");
    return adapter;
  }
  sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> { return this.worker(input.providerSessionId).sendTurn(input); }
  async cancelTurn(id: string) { await this.sessions.get(id)?.cancelTurn(id); }
  sessionOptions(id: string) { return this.sessions.get(id)?.sessionOptions?.(id) ?? []; }
  async setSessionOption(id: string, optionId: string, value: string) {
    const adapter = this.sessions.get(id);
    if (!adapter?.setSessionOption) throw new Error("This provider has no settings to change.");
    return adapter.setSessionOption(id, optionId, value);
  }
  async steerTurn(id: string, text: string): Promise<boolean> { return (await this.sessions.get(id)?.steerTurn?.(id, text)) ?? false; }
  async closeSession(id: string) { const adapter = this.sessions.get(id); this.sessions.delete(id); await adapter?.closeSession(id); }
  async dispose() { await Promise.all([...this.adapters.values()].map(adapter => adapter.dispose())); this.adapters.clear(); this.sessions.clear(); }
}
