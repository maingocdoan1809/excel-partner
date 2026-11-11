import path from "node:path";
import fs from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import type { ProviderId, ProviderEvent } from "../../protocol/src/index.js";
import type { ProviderAdapter, ProviderRuntime, SendTurnInput, StartSessionContext } from "./types.js";

export function providerPackageRoot(): string {
  let root = path.dirname(fileURLToPath(import.meta.url));
  for (let i = 0; i < 8; i++, root = path.dirname(root)) if (fs.existsSync(path.join(root, "package.json"))) return root;
  throw new Error("Provider package root not found");
}
const root = providerPackageRoot();
const store = createRequire(import.meta.url)(path.join(root, "bin", "runtime-store.cjs"));
type Profile = { id: string; mode: "managed" | "external"; command: string; adapterRoot: string; fallbackFrom?: string };

export function providerExecution(provider: ProviderId, env = process.env, login = false): { command: string; env: NodeJS.ProcessEnv } {
  const state = login ? store.readState(store.home(env)) : null;
  const profile = login && state.profiles[state.pending?.[provider]] || store.selectProfile(store.home(env), provider);
  return { command: profile?.command ?? env[provider.toUpperCase() + "_BIN"] ?? provider, env: store.childEnvironment(env, profile) };
}
export async function activateProviderAfterLogin(provider: ProviderId, env = process.env, canActivate = () => true) {
  const dataDir = store.home(env), state = store.readState(dataDir), profile = state.profiles[state.pending?.[provider]];
  if (!profile) return;
  const manager = await import(pathToFileURL(path.join(profile.adapterRoot, "bin", "runtime-manager.mjs")).href);
  const deployment = store.currentDeployment(dataDir);
  const activate = deployment.releases[deployment.active]?.root === profile.adapterRoot;
  await manager.installRuntime(profile.adapterRoot, provider, { dataDir, env, adapterRoot: profile.adapterRoot, activate, expectedPending: profile.id,
    canActivate: () => canActivate() && store.readState(dataDir).pending?.[provider] === profile.id });
}

/** Route each worker to a fixed adapter/SDK/CLI profile. Never retry a sent turn. */
export class ProfiledAdapter implements ProviderAdapter {
  private readonly sessions = new Map<string, { adapter: ProviderAdapter; profile: Profile | null }>();
  private readonly adapters = new Map<string, ProviderAdapter>();
  private readonly dataDir: string;
  constructor(readonly id: ProviderId, private readonly runtime: ProviderRuntime,
    private readonly factory: (runtime: ProviderRuntime) => ProviderAdapter) { this.dataDir = store.home(runtime.env ?? process.env); }
  runtimeRevision() { const state = store.readState(this.dataDir); return state.disabled?.includes(this.id) ? "disabled" : state.active[this.id] ?? state.legacy[this.id] ?? "legacy"; }
  private async resolve(nativeId?: string, override?: Profile): Promise<{ adapter: ProviderAdapter; profile: Profile | null }> {
    const profile: Profile | null = override ?? store.selectProfile(this.dataDir, this.id, nativeId);
    const key = profile?.id ?? "legacy";
    let adapter = this.adapters.get(key);
    if (!adapter) {
      const runtime = { ...this.runtime, env: store.childEnvironment(this.runtime.env ?? process.env, profile) };
      if (profile && store.inside(path.join(this.dataDir, "releases"), profile.adapterRoot) && path.resolve(profile.adapterRoot) !== root) {
        const entry = path.join(profile.adapterRoot, "dist", "providers", "src", `${this.id}.js`);
        if (!fs.existsSync(entry)) throw new Error("Saved provider adapter is missing. Repair the recorded release before resuming.");
        const module = await import(pathToFileURL(entry).href);
        const name = { claude: "ClaudeAdapter", copilot: "CopilotAdapter", codex: "CodexAdapter", opencode: "OpencodeAdapter" }[this.id];
        adapter = new module[name](runtime) as ProviderAdapter;
      } else adapter = this.factory(runtime);
      this.adapters.set(key, adapter);
    }
    return { adapter, profile };
  }
  async detect(force = false) {
    try { return await (await this.resolve()).adapter.detect(force); }
    catch (error) {
      const disabled = error instanceof Error && error.message === "PROVIDER_DISABLED";
      return { provider: this.id, installed: !disabled, version: null, status: disabled ? "not_installed" as const : "error" as const,
        hint: disabled ? `Install this provider with office-partner providers install ${this.id}.` : "Saved provider runtime is unavailable. Run office-partner doctor and office-partner repair.",
        capabilities: { streaming: true, cancel: true, sessionResume: true, mcp: true } };
    }
  }
  async connectionStatus() { return (await this.resolve()).adapter.connectionStatus(); }
  async listModels() { return (await this.resolve()).adapter.listModels?.() ?? { models: [], defaultModelId: null }; }
  async listCommands(id?: string) {
    const state = id ? this.sessions.get(id) : undefined;
    return (state?.adapter ?? (await this.resolve()).adapter).listCommands?.(id) ?? [];
  }
  async sessionTitle(nativeId: string) { return (await this.resolve(nativeId)).adapter.sessionTitle?.(nativeId); }
  async startSession(ctx: StartSessionContext) {
    let state = await this.resolve(ctx.resumeSessionId);
    let id: string;
    const open = async (adapter: ProviderAdapter) => {
      let starting = true, closed = false, accepted = false;
      const worker = await adapter.startSession({ ...ctx, onRuntimeClosed: () => {
        if (starting) closed = true;
        else if (accepted) ctx.onRuntimeClosed?.();
      } });
      starting = false;
      if (closed) { await adapter.closeSession(worker); throw new Error("Provider runtime closed during initialization"); }
      accepted = true;
      return worker;
    };
    try { id = await open(state.adapter); }
    catch (error) {
      // No prompt has been sent. Auth/policy/model errors never take this path.
      if (ctx.resumeSessionId || !state.profile || !/ENOENT|MODULE_NOT_FOUND|Cannot find module|ERR_DLOPEN_FAILED|CLI.*not found|unsupported protocol|protocol version mismatch|unknown option|unrecognized option/i.test(String(error))) throw error;
      const fallback = store.fallbackProfile(this.dataDir, this.id, state.profile.id);
      if (!fallback) throw error;
      state = await this.resolve(undefined, { ...fallback, fallbackFrom: state.profile.id });
      id = await open(state.adapter);
    }
    this.sessions.set(id, state);
    if (state.profile?.fallbackFrom) this.runtime.logger.warn("provider_runtime_fallback", { provider: this.id });
    return id;
  }
  async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const state = this.sessions.get(input.providerSessionId);
    if (!state) throw new Error("Provider worker is not available");
    for await (const event of state.adapter.sendTurn(input)) {
      if (event.type === "session_info" && state.profile) {
        for (let attempt = 0; ; attempt++) {
          try { store.bindConversation(this.dataDir, this.id, event.providerSessionId, state.profile.id); break; }
          catch (error) {
            if (attempt >= 9 || !(error instanceof Error) || error.message !== "OPERATION_LOCKED") throw error;
            await new Promise(resolve => setTimeout(resolve, 10 * (attempt + 1)));
          }
        }
      }
      yield event;
    }
  }
  async cancelTurn(id: string) { await this.sessions.get(id)?.adapter.cancelTurn(id); }
  async closeSession(id: string) { const state = this.sessions.get(id); this.sessions.delete(id); await state?.adapter.closeSession(id); }
  async dispose() { await Promise.all([...this.adapters.values()].map(adapter => adapter.dispose())); this.adapters.clear(); this.sessions.clear(); }
}
