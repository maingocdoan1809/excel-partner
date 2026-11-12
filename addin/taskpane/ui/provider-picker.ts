import { ProviderIdSchema, type ProviderConnectionStatus, type ProviderDetection, type ProviderId } from "../../../protocol/src/index.js";
import type { BridgeClient, ModelCatalog } from "../bridge-client.js";
import { setupComposerMenu } from "./composer-menu.js";
import { byId } from "./dom.js";
import type { StoredMessage } from "./sessions.js";

export const PROVIDER_KEY = "excelPartner.provider";
export const PROVIDER_NAMES: Record<ProviderId, string> = { claude: "Claude Code", copilot: "GitHub Copilot", codex: "Codex", opencode: "OpenCode" };
const MODEL_KEY = "excelPartner.models.v1";
/** Short label used in the model picker. */
const PROVIDER_SHORT_NAMES: Record<ProviderId, string> = { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode" };
const MODEL_ID = /^[a-zA-Z0-9][a-zA-Z0-9._/-]{0,79}$/;

function loadModels(): Partial<Record<ProviderId, string>> {
  try {
    const parsed = JSON.parse(localStorage.getItem(MODEL_KEY) ?? "{}") as Record<string, unknown>;
    return Object.fromEntries(Object.entries(parsed).filter(([key, value]) =>
      ProviderIdSchema.safeParse(key).success && typeof value === "string" && MODEL_ID.test(value),
    )) as Partial<Record<ProviderId, string>>;
  } catch { return {}; }
}

export function loadProvider(): ProviderId {
  const saved = localStorage.getItem(PROVIDER_KEY);
  return ProviderIdSchema.safeParse(saved).data ?? "copilot";
}

export interface ProviderPickerDeps {
  getClient(): BridgeClient;
  getProvider(): ProviderId;
  setProvider(provider: ProviderId): void;
  closeMenu(): void;
  openPanel(): string | null;
  startNewChat(keepPanel: boolean): void;
  refreshTools(): void;
  append(message: StoredMessage): HTMLElement;
  elements: { providerHint: HTMLElement; modelSelect: HTMLSelectElement };
}

/** Provider availability and the CLI's model catalog, with saved choices per provider. */
export function setupProviderPicker(deps: ProviderPickerDeps) {
  const els = deps.elements;
  const models = loadModels();
  const modelCatalogs = new Map<ProviderId, ModelCatalog>();
  const modelErrors = new Map<ProviderId, boolean>();
  const providerHints = new Map<ProviderId, string>();
  const providerStatuses = new Map<ProviderId, ProviderConnectionStatus>();
  const modelLoginRequired = new Set<ProviderId>();
  let modelPicker: ReturnType<typeof setupComposerMenu> | undefined;

  function renderProviderMenu(): void {
    deps.refreshTools();
    byId<HTMLSelectElement>("provider-select").value = deps.getProvider();
    els.providerHint.textContent = providerHints.get(deps.getProvider()) ?? "";
    renderAccount();
  }

  /** Says whether the chosen provider can be used now; the sign-in button only offers to switch accounts once it is. */
  function renderAccount(): void {
    const provider = deps.getProvider();
    const status = providerStatuses.get(provider) ?? "unknown";
    const name = PROVIDER_NAMES[provider];
    const line = byId("provider-status");
    line.dataset.state = status;
    line.textContent = status === "ready" ? `✓ Signed in to ${name} and ready to use`
      : status === "login_required" ? `Not signed in to ${name}`
      : status === "not_installed" ? `${name} is not installed on this computer`
      : status === "error" ? `${name} has a problem. See the hint above`
      : `Checking ${name}…`;
    byId("menu-login").textContent = status === "ready" ? `Sign in with another ${name} account`
      : status === "not_installed" ? `Set up ${name}`
      : `Sign in to ${name}`;
  }

  function statusOf(provider: ProviderId): ProviderConnectionStatus {
    return providerStatuses.get(provider) ?? "unknown";
  }

  function renderAuthNotice(): void {
    const status = providerStatuses.get(deps.getProvider());
    byId("auth-notice").hidden = status !== "login_required" && status !== "not_installed";
    byId("auth-notice-title").textContent = status === "not_installed" ? `Set up ${PROVIDER_NAMES[deps.getProvider()]}` : `Sign in to ${PROVIDER_NAMES[deps.getProvider()]}`;
    byId("auth-notice-detail").textContent = status === "not_installed" ? "Install the CLI on the bridge machine." : "Connect your account to start chatting.";
    byId("provider-login-button").textContent = status === "not_installed" ? "Set up ↗" : "Sign in ↗";
    renderAccount();
  }

  function selectProvider(next: ProviderId): void {
    deps.closeMenu();
    if (next === deps.getProvider()) return;
    deps.setProvider(next);
    localStorage.setItem(PROVIDER_KEY, next);
    renderProviderMenu();
    renderAuthNotice();
    renderModelPicker();
    // A provider keeps its own conversation: switching starts a fresh chat (the old one stays in History).
    deps.startNewChat(deps.openPanel() === "settings");
    deps.getClient().setProvider(next);
    void refreshModelCatalog(next);
  }

  function renderModelPicker(): void {
    const catalog = modelCatalogs.get(deps.getProvider());
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "Default";
    els.modelSelect.replaceChildren(defaultOption, ...(catalog?.models ?? []).map((model) => {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.name;
      return option;
    }));
    if (!catalog) {
      const info = document.createElement("option");
      info.disabled = true;
      info.textContent = modelErrors.get(deps.getProvider()) ? "Models unavailable" : "Loading models…";
      els.modelSelect.append(info);
    }
    const saved = models[deps.getProvider()];
    els.modelSelect.value = saved && catalog?.models.some((model) => model.id === saved) ? saved : "";
    els.modelSelect.title = `${PROVIDER_NAMES[deps.getProvider()]} model${catalog?.defaultModelId ? `; default: ${catalog.defaultModelId}` : ""}`;
    modelPicker?.refresh();
  }

  async function refreshModelCatalog(forProvider: ProviderId): Promise<void> {
    try {
      const catalog = await deps.getClient().listModels(forProvider);
      modelLoginRequired.delete(forProvider);
      if (providerStatuses.get(forProvider) === "login_required") providerStatuses.set(forProvider, "ready");
      renderAuthNotice();
      modelCatalogs.set(forProvider, catalog);
      modelErrors.delete(forProvider);
      if (deps.getProvider() === forProvider) renderModelPicker();
    } catch (error) {
      modelErrors.set(forProvider, true);
      if ((error as { code?: string }).code === "login_required") {
        modelLoginRequired.add(forProvider);
        providerStatuses.set(forProvider, "login_required");
        renderAuthNotice();
      }
      if (deps.getProvider() === forProvider) renderModelPicker();
    }
  }

  function saveModelChoice(): void {
    const model = els.modelSelect.value;
    if (model) models[deps.getProvider()] = model;
    else delete models[deps.getProvider()];
    localStorage.setItem(MODEL_KEY, JSON.stringify(models));
  }

  function setupModelPicker(): void {
    modelPicker = setupComposerMenu(els.modelSelect, byId<HTMLButtonElement>("model-button"), byId("model-menu"), (label) => `${PROVIDER_SHORT_NAMES[deps.getProvider()]} · ${label}`);
  }

  function requireLogin(provider: ProviderId): void {
    modelLoginRequired.add(provider);
    providerStatuses.set(provider, "login_required");
    renderAuthNotice();
  }

  function signedIn(provider: ProviderId): void {
    modelLoginRequired.delete(provider);
    providerStatuses.set(provider, "ready");
    renderAuthNotice();
  }

  function invalidateCatalog(provider: ProviderId): void {
    modelCatalogs.delete(provider);
  }

  function onStatus(id: ProviderId, detection: ProviderDetection): void {
    providerStatuses.set(id, detection.status === "ready" && modelLoginRequired.has(id) ? "login_required" : detection.status);
    providerHints.set(id, `${detection.version ?? "not found"} · ${detection.status.replace("_", " ")}`);
    renderProviderMenu();
    renderAuthNotice();
    if (id === deps.getProvider() && detection.status === "error") {
      deps.append({ kind: "error", text: detection.hint ?? `${PROVIDER_NAMES[id]} is ${detection.status}.` });
    }
  }

  return { statusOf, renderProviderMenu, renderAuthNotice, selectProvider, renderModelPicker, refreshModelCatalog, saveModelChoice,
    setupModelPicker, requireLogin, signedIn, invalidateCatalog, onStatus };
}
