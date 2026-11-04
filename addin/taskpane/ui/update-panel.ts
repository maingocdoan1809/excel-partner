/// <reference lib="dom" />

/** What the bridge reports about versions and the update in progress (GET /session/update). */
export interface UpdateInfo {
  current: string;
  latest: string | null;
  updateAvailable: boolean;
  canUpdate: boolean;
  reason?: string;
  checkError?: string;
  status: { state: "idle" | "running" | "done" | "failed"; step?: string; message?: string; from?: string | null; to?: string | null; error?: string | null };
  logFile: string;
}

export interface UpdateClient {
  updateInfo(refresh?: boolean): Promise<UpdateInfo>;
  startUpdate(): Promise<void>;
}

export interface UpdatePanelDeps {
  getClient(): UpdateClient;
  /** Reload the task pane so it runs the new files. */
  reload(): void;
  /** Where the "an update was in progress" marker lives; null disables it. */
  storage?: Pick<Storage, "getItem" | "setItem" | "removeItem"> | null;
  /** Called with a sentence once an update that was started here has finished, after the pane reloaded. */
  announce?(text: string): void;
  pollMs?: number;
  /** Give up waiting for the bridge to come back after this long. */
  waitMs?: number;
  confirmMs?: number;
}

const MARKER = "excelPartner.updating";

/** Settings > Updates: show the version, look for a newer one, and replace the installation without closing Excel. */
export function setupUpdatePanel(deps: UpdatePanelDeps) {
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const version = $("update-version");
  const status = $("update-status");
  const check = $<HTMLButtonElement>("update-check");
  const run = $<HTMLButtonElement>("update-run");
  const pollMs = deps.pollMs ?? 2000;
  const waitMs = deps.waitMs ?? 5 * 60_000;
  const confirmMs = deps.confirmMs ?? 6000;
  const storage = deps.storage === undefined ? safeStorage() : deps.storage;

  let polling: number | null = null;
  let confirming: number | null = null;
  let updating = false;
  let startedFrom = "";
  let lastInfo: UpdateInfo | null = null;

  const say = (text: string, kind: "info" | "error" | "ok" = "info") => {
    status.textContent = text;
    status.dataset.kind = kind;
  };
  const mark = (on: boolean) => { try { if (on) storage?.setItem(MARKER, startedFrom || "1"); else storage?.removeItem(MARKER); } catch { /* storage may be blocked */ } };

  function stopConfirming(): void {
    if (confirming !== null) window.clearTimeout(confirming);
    confirming = null;
    run.dataset.confirm = "false";
    run.textContent = lastInfo?.latest ? `Update to ${lastInfo.latest}` : "Update now";
  }

  function render(info: UpdateInfo): void {
    lastInfo = info;
    version.textContent = `Version ${info.current}`;
    check.disabled = updating;
    run.hidden = !(info.updateAvailable && info.canUpdate) || updating;
    if (!run.hidden) stopConfirming();
    if (updating) return;
    if (info.status.state === "failed") {
      say(`The last update did not finish: ${info.status.error ?? "unknown error"}. Details: ${info.logFile}`, "error");
    } else if (info.updateAvailable && !info.canUpdate) {
      say(`Version ${info.latest} is available. ${info.reason ?? ""}`.trim());
    } else if (info.updateAvailable) {
      say(`Version ${info.latest} is available. Excel stays open: the connection restarts for about a minute and this pane reloads by itself.`);
    } else if (info.latest) {
      say("You have the latest version.", "ok");
    } else {
      say(info.checkError ? `Could not check for updates: ${info.checkError}` : "Could not check for updates.", "error");
    }
  }

  async function refresh(force = false): Promise<void> {
    check.disabled = true;
    say("Checking for updates…");
    try {
      const info = await deps.getClient().updateInfo(force);
      check.disabled = false;
      render(info);
      if (info.status.state === "running" && !updating) resume(info);
    } catch (error) {
      check.disabled = false;
      say(`Could not check for updates: ${error instanceof Error ? error.message : String(error)}`, "error");
    }
  }

  function resume(info: UpdateInfo): void {
    updating = true;
    startedFrom = info.status.from ?? info.current;
    track(info.status.message ?? "Updating");
  }

  /** Follow the update until the new bridge answers; the old one goes away halfway, so errors are expected. */
  function track(initial: string): void {
    updating = true;
    check.disabled = true;
    run.hidden = true;
    say(`${initial}… Keep Excel open.`);
    const started = Date.now();
    if (polling !== null) window.clearInterval(polling);
    polling = window.setInterval(async () => {
      let info: UpdateInfo;
      try {
        info = await deps.getClient().updateInfo(false);
      } catch {
        // The bridge is down while it is replaced.
        if (Date.now() - started > waitMs) finish("The bridge did not come back. Run excel-partner start in a terminal, or see update.log in the Excel Partner folder.", "error");
        else say("Restarting the connection… Keep Excel open.");
        return;
      }
      const s = info.status;
      if (s.state === "running") say(`${s.message ?? "Updating"}… Keep Excel open.`);
      else if (s.state === "done") {
        // The marker stays: the reloaded pane reads it and says in the chat which version it is now on.
        finish(`Updated to ${s.to ?? info.current}. Reloading…`, "ok");
        window.setTimeout(() => deps.reload(), 1200);
      } else if (s.state === "failed") {
        finish(`The update did not finish: ${s.error ?? "unknown error"}. Details: ${info.logFile}`, "error");
        mark(false);
        lastInfo = info;
      } else if (Date.now() - started > 30_000) {
        finish("The update did not start. Try again.", "error");
        mark(false);
      }
    }, pollMs);
  }

  function finish(text: string, kind: "info" | "error" | "ok"): void {
    if (polling !== null) window.clearInterval(polling);
    polling = null;
    updating = false;
    check.disabled = false;
    say(text, kind);
  }

  check.addEventListener("click", () => void refresh(true));
  run.addEventListener("click", async () => {
    if (!lastInfo || updating) return;
    // Two clicks: starting replaces the connection, so a stray click should not do it.
    if (confirming === null) {
      run.dataset.confirm = "true";
      run.textContent = "Click again to update";
      confirming = window.setTimeout(stopConfirming, confirmMs);
      return;
    }
    stopConfirming();
    startedFrom = lastInfo.current;
    run.disabled = true;
    try {
      await deps.getClient().startUpdate();
      mark(true);
      track("Starting the update");
    } catch (error) {
      say(`Could not start the update: ${error instanceof Error ? error.message : String(error)}`, "error");
    } finally {
      run.disabled = false;
    }
  });

  /** After a reload: if this pane started an update, say how it ended. */
  async function afterReload(): Promise<void> {
    let from: string | null = null;
    try { from = storage?.getItem(MARKER) ?? null; } catch { from = null; }
    if (!from) return;
    try {
      const info = await deps.getClient().updateInfo(false);
      if (info.status.state === "running") return; // still going; Settings will pick it up
      mark(false);
      if (info.status.state === "done") deps.announce?.(`Excel Partner was updated to version ${info.status.to ?? info.current}.`);
      else if (info.status.state === "failed") deps.announce?.(`The update did not finish: ${info.status.error ?? "unknown error"}. Details: ${info.logFile}`);
    } catch { /* bridge not ready yet; the marker stays for the next load */ }
  }

  return { refresh: () => refresh(false), afterReload };
}

function safeStorage(): Storage | null {
  try { return window.localStorage; } catch { return null; }
}
