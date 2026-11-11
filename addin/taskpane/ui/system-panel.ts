import type { ProviderId } from "../../../protocol/src/index.js";
import { confirmAction } from "./confirm-action.js";

export interface SystemInfo {
  version: string | null; held: boolean; phase: string | null; debug?: boolean;
  providers: Array<{ provider: ProviderId; selected: boolean; disabled?: boolean; mode: string; version: string | null; pendingVersion: string | null; pendingVerified?: boolean; available: boolean | null }>;
}
export interface DiagnosticReport { status: string; diagnosticId: string; checks: Array<{ status: string; id: string; code: string; message: string }> }
interface SystemClient {
  systemInfo(): Promise<SystemInfo>;
  systemAction(action: string, provider?: ProviderId, version?: string): Promise<void>;
  systemDiagnostics(options?: { provider?: ProviderId; probe?: boolean; mcp?: boolean }): Promise<DiagnosticReport>;
  updateInfo(refresh?: boolean): Promise<{ status: { state: string; error?: string | null } }>;
}
export function setupSystemPanel(getClient: () => SystemClient, isBusy: () => boolean) {
  const section = document.getElementById("settings-system-section");
  if (!section) return { refresh: async () => {} };
  const list = document.getElementById("system-providers")!;
  const status = document.getElementById("system-status")!;
  const output = document.getElementById("system-report")!;
  const summary = document.getElementById("system-version")!;
  const exportButton = document.getElementById("system-export")!, debugButton = document.getElementById("system-debug")!;
  let working = false, debugOn = false, report: DiagnosticReport | undefined;
  const say = (message: string) => { status.textContent = message; };
  const disable = (value: boolean) => { working = value; for (const button of section.querySelectorAll<HTMLButtonElement>("button")) button.disabled = value; };
  async function operation(action: () => Promise<void>) {
    if (working) return;
    if (isBusy()) { say("Wait for the current response to finish."); return; }
    disable(true);
    try { await action(); }
    catch (error) { say(error instanceof Error ? error.message : "Operation failed. Export diagnostics before retrying."); }
    finally { disable(false); }
  }
  async function refresh() {
    try {
      const info = await getClient().systemInfo();
      summary.textContent = `Installed: ${info.version ?? "current installation"}.${info.held ? " Automatic activation is held after recovery." : ""}`;
      debugOn = info.debug === true; debugButton.textContent = debugOn ? "Stop debug" : "Debug for 15 min";
      list.replaceChildren();
      for (const provider of info.providers) {
        const row = document.createElement("div"); row.className = "system-provider";
        const copy = document.createElement("div"); copy.className = "system-provider-copy";
        const name = document.createElement("strong");
        name.textContent = { claude: "Claude Code", copilot: "GitHub Copilot", codex: "Codex", opencode: "OpenCode" }[provider.provider];
        const text = document.createElement("p"); text.className = "settings-description";
        text.textContent = `${provider.disabled ? "Disabled" : provider.mode === "managed" ? `private CLI ${provider.version}${provider.available === false ? " (repair required)" : ""}` : "Using your installed CLI"}${provider.pendingVersion ? `; ${provider.pendingVersion} ${provider.pendingVerified ? "prepared for release switch" : "awaiting verification/sign-in"}` : ""}`;
        const button = document.createElement("button"); button.type = "button"; button.className = "secondary";
        button.textContent = provider.selected ? "Disable" : provider.pendingVersion ? "Verify again" : "Install private CLI";
        button.addEventListener("click", () => void operation(async () => {
          say(provider.selected ? "Disabling provider…" : "Installing and verifying the private CLI. This may take a few minutes…");
          try { await getClient().systemAction(provider.selected ? "remove" : "install", provider.provider); say("Provider configuration saved. Existing chats keep their runtime."); }
          finally { await refresh(); }
        }));
        copy.append(name, text); row.append(copy, button); list.append(row);
        if (provider.pendingVersion && !provider.pendingVerified) { const hint = document.createElement("p"); hint.className = "settings-description"; hint.textContent = `Sign in with: office-partner providers login ${provider.provider}${provider.provider === "copilot" ? ". You can also use Sign in in Settings." : ""}`; copy.append(hint); }
      }
    } catch { say("System information is unavailable. Run office-partner-doctor if the bridge is offline."); }
  }
  async function check() {
    delete status.dataset.state;
    say("Checking the installation and the local MCP schemas…");
    report = await getClient().systemDiagnostics({ mcp: true });
    exportButton.hidden = false;
    output.textContent = report.checks.map(check => `[${check.status}] ${check.id} · ${check.code}: ${check.message}`).join("\n");
    const failed = report.checks.filter(check => check.status === "fail");
    status.dataset.state = report.status;
    say(`${failed.length ? `${failed.length} check${failed.length === 1 ? "" : "s"} failed: ${failed.map(check => check.id).join(", ")}` : `Check finished: ${report.status}`}. Reference: ${report.diagnosticId}`);
  }
  document.getElementById("system-check")!.addEventListener("click", () => void operation(() => check()));
  exportButton.addEventListener("click", () => void operation(async () => {
    if (!report) return;
    const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }));
    const link = document.createElement("a"); link.href = url; link.download = "office-partner-diagnostics.json"; link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000); say("Diagnostic metadata exported. No raw logs or document contents were included.");
  }));
  async function repair() {
    if (!await confirmAction(section!, "Check the installed files and repair them and the private CLIs? A damaged installation is installed again from npm. Office Partner will reconnect.", "Repair installation")) return;
    await getClient().systemAction("repair");
    say("Recovery started. Waiting for Office Partner to reconnect…");
    for (let i = 0; i < 180; i++) {
      await new Promise(resolve => window.setTimeout(resolve, 2000));
      try {
        const result = await getClient().updateInfo();
        if (result.status.state === "done") { say("Recovery completed."); await refresh(); return; }
        if (result.status.state === "failed") { say(result.status.error ?? "Recovery failed. Use the standalone recovery launcher."); return; }
      } catch { /* service is restarting */ }
    }
    say("The service has not reconnected. Use office-partner-doctor or the recovery launcher.");
  }
  document.getElementById("system-repair")!.addEventListener("click", () => void operation(repair));
  debugButton.addEventListener("click", () => void operation(async () => {
    const turnOn = !debugOn;
    await getClient().systemAction(turnOn ? "debug-on" : "debug-off");
    debugOn = turnOn; debugButton.textContent = turnOn ? "Stop debug" : "Debug for 15 min";
    say(turnOn ? "Debug logging enabled for 15 minutes." : "Debug logging disabled.");
  }));
  return { refresh };
}
