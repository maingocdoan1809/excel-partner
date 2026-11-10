import type { ProviderId } from "../../../protocol/src/index.js";

export interface SystemInfo {
  version: string | null; previous: string | null; held: boolean; phase: string | null; cached: string[];
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
  const versions = document.getElementById("system-rollback-version") as HTMLSelectElement;
  const summary = document.getElementById("system-version")!;
  let working = false, report: DiagnosticReport | undefined;
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
      summary.textContent = `Active: ${info.version ?? "current installation"}. Previous: ${info.previous ?? "none retained yet"}.${info.held ? " Automatic activation is held after recovery." : ""}`;
      versions.replaceChildren();
      for (const version of [...new Set(info.cached ?? [])]) { const option = document.createElement("option"); option.value = version; option.textContent = version; versions.append(option); }
      list.replaceChildren();
      for (const provider of info.providers) {
        const row = document.createElement("div"); row.className = "settings-row";
        const text = document.createElement("p"); text.className = "settings-description";
        text.textContent = `${provider.provider}: ${provider.disabled ? "disabled" : provider.mode === "managed" ? `private CLI ${provider.version}${provider.available === false ? " (repair required)" : ""}` : "existing system CLI"}${provider.pendingVersion ? `; ${provider.pendingVersion} ${provider.pendingVerified ? "prepared for release switch" : "awaiting verification/sign-in"}` : ""}`;
        const button = document.createElement("button"); button.type = "button"; button.className = "secondary";
        button.textContent = provider.selected ? "Disable" : provider.pendingVersion ? "Verify again" : "Install private CLI";
        button.addEventListener("click", () => void operation(async () => {
          say(provider.selected ? "Disabling provider…" : "Installing and verifying the private CLI. This may take a few minutes…");
          try { await getClient().systemAction(provider.selected ? "remove" : "install", provider.provider); say("Provider configuration saved. Existing chats keep their runtime."); }
          finally { await refresh(); }
        }));
        row.append(text, button); list.append(row);
        if (provider.pendingVersion && !provider.pendingVerified) { const hint = document.createElement("p"); hint.className = "settings-description"; hint.textContent = `Sign in with: office-partner providers login ${provider.provider}${provider.provider === "copilot" ? ". You can also use Sign in in Settings." : ""}`; list.append(hint); }
      }
    } catch { say("System information is unavailable. Run office-partner-doctor if the bridge is offline."); }
  }
  async function check(mcp = false) {
    say(mcp ? "Checking local MCP schemas…" : "Checking the installation…");
    report = await getClient().systemDiagnostics({ mcp });
    output.textContent = report.checks.map(check => `[${check.status}] ${check.code}: ${check.message}`).join("\n");
    say(`Check finished: ${report.status}. Reference: ${report.diagnosticId}`);
  }
  document.getElementById("system-check")!.addEventListener("click", () => void operation(() => check()));
  document.getElementById("system-check-mcp")!.addEventListener("click", () => void operation(() => check(true)));
  document.getElementById("system-export")!.addEventListener("click", () => void operation(async () => {
    if (!report) await check();
    const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }));
    const link = document.createElement("a"); link.href = url; link.download = "office-partner-diagnostics.json"; link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000); say("Diagnostic metadata exported. No raw logs or document contents were included.");
  }));
  async function recover(action: "repair" | "rollback") {
    if (action === "rollback" && !versions.value) { say("No retained version is available yet."); return; }
    if (!window.confirm(action === "rollback" ? `Return to Office Partner ${versions.value}? Existing document edits are kept.` : "Verify the release and repair its private CLIs? Office Partner will reconnect.")) return;
    await getClient().systemAction(action, undefined, action === "rollback" ? versions.value : undefined);
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
  document.getElementById("system-repair")!.addEventListener("click", () => void operation(() => recover("repair")));
  document.getElementById("system-rollback")!.addEventListener("click", () => void operation(() => recover("rollback")));
  for (const state of ["on", "off"]) document.getElementById(`system-debug-${state}`)!.addEventListener("click", () => void operation(async () => {
    await getClient().systemAction(`debug-${state}`); say(state === "on" ? "Debug logging enabled for 15 minutes." : "Debug logging disabled.");
  }));
  return { refresh };
}
