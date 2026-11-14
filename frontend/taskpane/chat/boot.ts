import { activeHost, initHost } from "../host/index.js";
import { recoverCalculationMode } from "../executor/bulk.js";
import { applyHostBranding } from "../ui/host-branding.js";
import { HOST_BASELINE, isSupported } from "../host/requirements.js";
import { showToast } from "../ui/toast.js";
import { setupTheme } from "../ui/theme.js";
import { byId } from "../ui/dom.js";
import { setupComposerMenu } from "../ui/composer-menu.js";
import { setupSelectionChip } from "../ui/selection-chip.js";
import { setupInputResizer } from "../ui/input-resizer.js";
import { takeChatRestore } from "../ui/chat-restore.js";
import { createBridgeClient } from "./bridge-handlers.js";
import { bindListeners } from "./listeners.js";
import { providerPicker, refreshUsage, renderChanges, renderModelPicker, renderProviderMenu, store, updater } from "./services.js";
import { els } from "../ui/elements.js";
import { loadSession } from "./chats.js";
import { renderWelcome } from "./messages.js";
import { app } from "./state.js";

const NARROW_BELOW = 330;
/** After the pane reloaded itself, a chat younger than this is reopened automatically. */
const RESTORE_WITHIN_MS = 12 * 3_600_000;

/** Runs once Office is ready: checks the host, connects to the bridge, hooks up the pane and brings the last chat back. */
export function boot(host: string): void {
  if (!["Excel", "Word", "PowerPoint"].includes(host)) return;
  initHost(host);
  // A previous pane that stopped mid-write may have left calculation on manual.
  if (activeHost().id === "excel") void recoverCalculationMode().catch(() => undefined);
  applyHostBranding();
  {
    // The manifest only asks for what every host has; what this host needs is checked here, so an old Office says so.
    const [set, version] = HOST_BASELINE[activeHost().id];
    if (!isSupported([set, version])) window.setTimeout(() => showToast(`Office Partner needs ${set} ${version} in ${activeHost().app}, which this version of Office does not have. Update Office to use it.`), 0);
  }

  app.client = createBridgeClient();
  app.client.setPermissionMode(app.permissionMode);
  app.client.connect();

  setupTheme(byId<HTMLSelectElement>("theme-select"));
  providerPicker.setupModelPicker();
  const savedEffort = localStorage.getItem("excelPartner.effort");
  els.effortSelect.value = savedEffort === "low" || savedEffort === "high" ? savedEffort : "medium";
  const effortButton = byId<HTMLButtonElement>("effort-button");
  const refreshEffort = () => {
    const effort = els.effortSelect.value;
    const label = els.effortSelect.selectedOptions[0]!.textContent;
    effortButton.dataset.effort = effort;
    effortButton.setAttribute("aria-label", `Reasoning effort: ${label}`);
    els.effortSelect.title = `Reasoning effort: ${label}. Models without effort control use provider defaults. Changes apply to your next message.`;
    effortButton.title = els.effortSelect.title;
    localStorage.setItem("excelPartner.effort", effort);
  };
  refreshEffort();
  els.effortSelect.addEventListener("change", refreshEffort);
  setupComposerMenu(els.effortSelect, effortButton, byId("effort-menu"));
  app.permissionPicker = setupComposerMenu(els.permissionMode, byId<HTMLButtonElement>("permission-button"), byId("permission-menu"));

  app.selectionChip = setupSelectionChip(els.chip);
  updater.poll();
  setupInputResizer(els.resizer, els.input);
  const observer = new ResizeObserver(([entry]) => {
    document.body.dataset.narrow = String((entry?.contentRect.width ?? window.innerWidth) < NARROW_BELOW);
  });
  observer.observe(document.body);
  bindListeners();

  // Bring the last chat back only after the pane reloaded itself (a new build or a finished update). A pane that was
  // closed and opened again starts a new chat; the old one is in the history list.
  const restoring = takeChatRestore();
  try { app.restoreChatId = restoring ? sessionStorage.getItem("excelPartner.activeChatId") : null; } catch {}
  const previous = restoring ? (app.restoreChatId ? store.getActive(app.restoreChatId) : store.latest(RESTORE_WITHIN_MS)) : null;
  if (previous) { app.restoreChatId ??= previous.id; loadSession(previous); }
  else renderWelcome();
  renderProviderMenu();
  renderModelPicker();
  renderChanges();
  void refreshUsage();
}
