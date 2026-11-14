import type { ChatHistoryStore } from "./durable-history.js";
import type { ProviderId, TokenUsage, UsageSummary } from "../../../protocol/src/index.js";
import type { BridgeClient } from "../bridge-client.js";
import type { StoredMessage, StoredSession } from "./sessions.js";
import { formatCredits, formatDuration, usageCsv } from "./usage.js";

export interface UsagePanelDeps {
  getClient(): BridgeClient;
  getSession(): StoredSession;
  getProvider(): ProviderId;
  showsCredits(): boolean;
  store: ChatHistoryStore;
  persistNow(): void;
  append(message: StoredMessage): HTMLElement;
  elements: {
    usageStats: HTMLElement;
    usageBreakdown: HTMLElement;
    usageCacheSplit: HTMLElement;
    usageCredits: HTMLElement;
    usageSource: HTMLElement;
    usageExport: HTMLButtonElement;
  };
}

/** Accounting for the current chat, including the live turn until its persisted summary arrives. */
export function setupUsagePanel(deps: UsagePanelDeps) {
  const els = deps.elements;
  const store = deps.store;
  let usageSummary: UsageSummary | undefined;
  let usageSessionId: string | null = null;
  let usageRequestId: string | null = null;
  let usageStartedAt: number | null = null;
  let turnUsage: TokenUsage | null = null;
  let usageTimer: number | undefined;

  async function refreshUsage(): Promise<void> {
    const id = deps.getSession().id;
    if (usageSummary?.sessionId !== id) usageSummary = undefined;
    renderUsage();
    if (!deps.getClient()?.connected) return;
    try {
      const summaries = await deps.getClient().usageSummaries(id);
      if (deps.getSession().id !== id) return;
      usageSummary = summaries[0];
      if (usageSummary?.title) {
        deps.getSession().providerTitle = usageSummary.title;
        deps.getSession().title = usageSummary.title;
        store.save(deps.getSession());
      }
      renderUsage();
    } catch {
      if (deps.getSession().id === id) els.usageStats.textContent = "Usage statistics unavailable";
    }
  }

  function renderUsage(): void {
    const running = usageSessionId === deps.getSession().id && usageStartedAt !== null;
    const summary = usageSummary?.sessionId === deps.getSession().id ? usageSummary : undefined;
    const reported = (summary?.reportedTurns ?? 0) + (running && turnUsage ? 1 : 0);
    const turns = (summary?.turns ?? 0) + (running ? 1 : 0);
    const tokens = (summary?.totalTokens ?? 0) + (running ? turnUsage?.totalTokens ?? 0 : 0);
    const partial = reported > 0 && reported < turns;
    const duration = (summary?.durationMs ?? 0) + (running ? Date.now() - usageStartedAt! : 0);
    els.usageStats.textContent = `Session tokens: ${reported || !turns ? tokens.toLocaleString() : "N/A"}${partial ? " (partial)" : ""} · AI time: ${formatDuration(duration)}`;
    const input = (summary?.inputTokens ?? 0) + (running ? turnUsage?.inputTokens ?? 0 : 0);
    const cached = (summary?.cachedInputTokens ?? 0) + (running ? turnUsage?.cachedInputTokens ?? 0 : 0);
    const output = (summary?.outputTokens ?? 0) + (running ? turnUsage?.outputTokens ?? 0 : 0);
    els.usageBreakdown.textContent = reported || !turns
      ? `Input: ${input.toLocaleString()} · Cache read: ${cached.toLocaleString()} · Output: ${output.toLocaleString()}`
      : "Provider has not reported token usage";
    els.usageBreakdown.title = "Cached tokens are included in input, not added again. Input includes instructions, tools, workbook context and conversation history.";
    const splitTurns = (summary?.cacheSplitReportedTurns ?? 0) + (running && turnUsage?.cacheWriteTokens !== undefined && turnUsage?.uncachedInputTokens !== undefined ? 1 : 0);
    const uncached = (summary?.uncachedInputTokens ?? 0) + (running ? turnUsage?.uncachedInputTokens ?? 0 : 0);
    const cacheWrite = (summary?.cacheWriteTokens ?? 0) + (running ? turnUsage?.cacheWriteTokens ?? 0 : 0);
    els.usageCacheSplit.textContent = splitTurns
      ? `Uncached input: ${uncached.toLocaleString()} · Cache write: ${cacheWrite.toLocaleString()}${splitTurns < turns ? " (partial)" : ""}`
      : "Uncached input / cache write: not reported";
    els.usageCacheSplit.title = "Input = uncached input + cache read + cache write when the full split is available. Both cache buckets are already included in input.";
    const creditTurns = (summary?.creditReportedTurns ?? 0) + (running && turnUsage?.copilotCredits !== undefined ? 1 : 0);
    const credits = (summary?.copilotCredits ?? 0) + (running ? turnUsage?.copilotCredits ?? 0 : 0);
    const exactCreditTurns = (summary?.exactCreditReportedTurns ?? 0) + (running && turnUsage?.copilotCreditNanoAiu !== undefined ? 1 : 0);
    els.usageCredits.hidden = !deps.showsCredits() && creditTurns === 0;
    els.usageCredits.textContent = creditTurns
      ? `Copilot credits: ${exactCreditTurns === creditTurns ? "" : "≈ "}${formatCredits(credits)}${creditTurns < turns ? " (partial)" : ""}${exactCreditTurns < creditTurns ? " (includes rounded legacy data)" : ""}`
      : "Copilot credits: not reported";
    els.usageCredits.title = "Original Copilot credits retain integer nano AI units (1 credit = 1,000,000,000 nano units). Old /usage records remain marked rounded. Credits are never estimated from tokens.";
    const originalTurns = (summary?.originalReportedTurns ?? 0) + (running && turnUsage?.source && turnUsage.source !== "copilot_acp" ? 1 : 0);
    const cliSessions = new Set(summary?.providerSessionIds ?? []);
    if (running && turnUsage?.providerSessionId) cliSessions.add(turnUsage.providerSessionId);
    els.usageSource.hidden = cliSessions.size === 0;
    els.usageSource.textContent = `Original provider data: ${originalTurns}/${turns} turns · CLI sessions: ${cliSessions.size}`;
    els.usageSource.title = `Each native CLI session has separate history and accounting. Session: ${[...cliSessions].join(", ") || "not recorded"}. Resuming keeps its ID and previously recorded usage. Background calls after turn completion are outside this report.`;
    els.usageStats.title = `This chat session: ${reported}/${turns} turns with token data. Total includes cached tokens. AI time includes provider/tool/approval waits while processing; idle time is excluded.`;
  }

  async function exportUsage(): Promise<void> {
    els.usageExport.disabled = true;
    try {
      deps.persistNow();
      const summaries = await deps.getClient().usageSummaries();
      const titles = new Map(store.list().map((item) => [item.id, item.title]));
      const blob = new Blob([usageCsv(summaries, titles)], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `excel-partner-usage-${new Date().toISOString().slice(0, 10)}.csv`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      deps.append({ kind: "error", text: `Could not export usage: ${error instanceof Error ? error.message : String(error)}` });
    } finally {
      els.usageExport.disabled = false;
    }
  }

  function start(requestId: string): void {
    usageSessionId = deps.getSession().id;
    usageRequestId = requestId;
    usageStartedAt = Date.now();
    turnUsage = null;
    window.clearInterval(usageTimer);
    usageTimer = window.setInterval(renderUsage, 1000);
    renderUsage();
  }

  function update(requestId: string, usage: TokenUsage): void {
    if (requestId === usageRequestId) {
      turnUsage = usage;
      renderUsage();
    }
  }

  function reset(): void {
    window.clearInterval(usageTimer);
    usageStartedAt = null;
    turnUsage = null;
    usageRequestId = null;
  }

  function finish(requestId: string): void {
    if (requestId === usageRequestId) reset();
  }

  function rebind(requestId: string): void {
    usageSummary = undefined;
    if (usageRequestId === requestId) usageSessionId = deps.getSession().id;
  }

  return { refreshUsage, renderUsage, exportUsage, start, update, reset, finish, rebind };
}
