import type { ProviderAdapter, ProviderModelCatalog, ProviderRuntime } from "./types.js";
import type { ProviderId } from "../../protocol/src/index.js";
import { ClaudeAdapter } from "./claude.js";
import { CopilotAdapter } from "./copilot.js";
import { CodexAdapter } from "./codex.js";
import { OpencodeAdapter } from "./opencode.js";

export class ProviderRegistry {
  private readonly adapters = new Map<ProviderId, ProviderAdapter>();
  private readonly modelCatalogs = new Map<ProviderId, { at: number; value: ProviderModelCatalog }>();

  register(adapter: ProviderAdapter): this {
    this.adapters.set(adapter.id, adapter);
    this.modelCatalogs.delete(adapter.id);
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
    const cached = this.modelCatalogs.get(id);
    if (cached && Date.now() - cached.at < 5 * 60_000) return cached.value;
    const value = await this.require(id).listModels?.() ?? { models: [], defaultModelId: null };
    this.modelCatalogs.set(id, { at: Date.now(), value });
    return value;
  }

  invalidateModels(id: ProviderId): void {
    this.modelCatalogs.delete(id);
  }

  async detectAll() {
    return Promise.all([...this.adapters.values()].map((adapter) => adapter.detect()));
  }

  async disposeAll(): Promise<void> {
    await Promise.all([...this.adapters.values()].map((adapter) => adapter.dispose()));
    this.adapters.clear();
  }
}

export function createDefaultRegistry(runtime: ProviderRuntime): ProviderRegistry {
  return new ProviderRegistry()
    .register(new ClaudeAdapter(runtime))
    .register(new CopilotAdapter(runtime))
    .register(new CodexAdapter(runtime))
    .register(new OpencodeAdapter(runtime));
}
