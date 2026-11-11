import type { ProviderAdapter, ProviderModelCatalog, ProviderRuntime } from "./types.js";
import { PROVIDER_IDS, type ProviderId } from "../../protocol/src/index.js";
import { ClaudeAdapter } from "./claude.js";
import { CopilotAdapter } from "./copilot.js";
import { CodexAdapter } from "./codex.js";
import { OpencodeAdapter } from "./opencode.js";
import { ProfiledAdapter } from "./profiled.js";

export class ProviderRegistry {
  private readonly adapters = new Map<ProviderId, ProviderAdapter>();
  private readonly modelCatalogs = new Map<ProviderId, { at: number; value: ProviderModelCatalog; revision?: string }>();
  private readonly modelLoading = new Map<ProviderId, { revision?: string; adapter: ProviderAdapter; promise: Promise<ProviderModelCatalog> }>();

  register(adapter: ProviderAdapter): this {
    this.adapters.set(adapter.id, adapter);
    this.invalidateModels(adapter.id);
    return this;
  }

  get(id: ProviderId): ProviderAdapter | undefined {
    return this.adapters.get(id);
  }

  require(id: ProviderId): ProviderAdapter {
    const adapter = this.get(id);
    if (!adapter) throw new Error(`no adapter registered for provider '${id}'`);
    return adapter;
  }

  ids(): ProviderId[] {
    return [...this.adapters.keys()];
  }

  async listModels(id: ProviderId): Promise<ProviderModelCatalog> {
    const adapter = this.require(id), revision = adapter.runtimeRevision?.();
    const cached = this.modelCatalogs.get(id);
    if (cached && cached.revision === revision && Date.now() - cached.at < 5 * 60_000) return cached.value;
    const pending = this.modelLoading.get(id);
    if (pending && pending.revision === revision && pending.adapter === adapter) return pending.promise;
    const loading = { revision, adapter, promise: undefined as unknown as Promise<ProviderModelCatalog> };
    loading.promise = Promise.resolve().then(() => adapter.listModels?.() ?? { models: [], defaultModelId: null }).then(value => {
      if (this.modelLoading.get(id) === loading && this.get(id) === adapter && adapter.runtimeRevision?.() === revision) this.modelCatalogs.set(id, { at: Date.now(), value, revision });
      return value;
    }).finally(() => { if (this.modelLoading.get(id) === loading) this.modelLoading.delete(id); });
    this.modelLoading.set(id, loading);
    return loading.promise;
  }

  invalidateModels(id: ProviderId): void {
    this.modelCatalogs.delete(id);
    this.modelLoading.delete(id);
  }

  async detectAll() {
    return Promise.all([...this.adapters.values()].map((adapter) => adapter.detect()));
  }

  async disposeAll(): Promise<void> {
    await Promise.all([...this.adapters.values()].map((adapter) => adapter.dispose()));
    this.adapters.clear();
    this.modelCatalogs.clear(); this.modelLoading.clear();
  }
}

/** Exhaustive against the shared protocol catalog: adding an ID requires a factory. */
const factories: Record<ProviderId, (runtime: ProviderRuntime) => ProviderAdapter> = {
  claude: runtime => new ClaudeAdapter(runtime),
  copilot: runtime => new CopilotAdapter(runtime),
  codex: runtime => new CodexAdapter(runtime),
  opencode: runtime => new OpencodeAdapter(runtime),
};

export function createDefaultRegistry(runtime: ProviderRuntime): ProviderRegistry {
  const registry = new ProviderRegistry();
  for (const id of PROVIDER_IDS) registry.register(new ProfiledAdapter(id, runtime, factories[id]));
  return registry;
}
