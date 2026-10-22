import { providerSessionKey, type UsageSummary } from "../../../protocol/src/index.js";

export function formatDuration(ms: number): string {
  const seconds = Math.floor(Math.max(0, ms) / 1000);
  return seconds >= 3600
    ? `${Math.floor(seconds / 3600)}h ${Math.floor(seconds % 3600 / 60)}m ${seconds % 60}s`
    : `${Math.floor(seconds / 60)}m ${seconds % 60}s`;
}

/** Nine decimal places retain Copilot's nano AI units; trim display zeros only. */
export function formatCredits(credits: number): string {
  return credits.toFixed(9).replace(/\.?0+$/, "");
}

export function creditDataStatus(s: UsageSummary): string {
  if (!s.creditReportedTurns) return "unavailable";
  const coverage = s.creditReportedTurns === s.turns ? "complete" : "partial";
  return `${coverage} (${s.exactCreditReportedTurns === s.creditReportedTurns ? "original" : s.exactCreditReportedTurns ? "mixed original/rounded" : "rounded"})`;
}

/** BOM for Excel; quote fields and neutralise spreadsheet formulas in titles/IDs. */
export function usageCsv(summaries: UsageSummary[], titles: Map<string, string> = new Map()): string {
  const cell = (value: string | number) => {
    let text = String(value);
    if (typeof value === "string" && /^[\s]*[=+\-@]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g, '""')}"`;
  };
  const rows: Array<Array<string | number>> = [[
    "Session ID", "Title", "Started at (UTC)", "Last finished at (UTC)", "Turns", "Turns with token data",
    "AI processing seconds", "Average seconds per turn", "Input tokens (includes cache)", "Output tokens",
    "Cached input tokens (subset)", "Total tokens", "Token data status",
    "Copilot AI credits", "Turns with credit data", "Credit data status",
    "Uncached input tokens (subset, known turns)", "Cache write tokens (subset, known turns)", "Turns with cache split", "Cache split status",
    "Original Copilot credit nano AI units (known turns)", "Turns with original credits", "Turns with original token data", "CLI session IDs", "Provider",
    "Reasoning output tokens (subset, known turns)", "Turns with reasoning data", "Session grouping",
  ]];
  for (const s of summaries) rows.push([
    s.sessionId, s.title ?? titles.get(s.sessionId) ?? "", new Date(s.startedAt).toISOString(), new Date(s.updatedAt).toISOString(),
    s.turns, s.reportedTurns, (s.durationMs / 1000).toFixed(3), (s.durationMs / Math.max(1, s.turns) / 1000).toFixed(3),
    s.reportedTurns ? s.inputTokens : "", s.reportedTurns ? s.outputTokens : "",
    s.reportedTurns ? s.cachedInputTokens : "", s.reportedTurns ? s.totalTokens : "",
    s.reportedTurns === s.turns ? "complete" : s.reportedTurns ? "partial" : "unavailable",
    s.creditReportedTurns ? (s.exactCreditReportedTurns ? formatCredits(s.copilotCredits ?? 0) : (s.copilotCredits ?? 0).toFixed(2)) : "", s.creditReportedTurns ?? 0,
    creditDataStatus(s),
    s.cacheSplitReportedTurns ? s.uncachedInputTokens ?? "" : "", s.cacheSplitReportedTurns ? s.cacheWriteTokens ?? "" : "",
    s.cacheSplitReportedTurns ?? 0,
    s.cacheSplitReportedTurns === s.turns ? "complete" : s.cacheSplitReportedTurns ? "partial" : "unavailable",
    s.copilotCreditNanoAiu ?? "", s.exactCreditReportedTurns ?? 0, s.originalReportedTurns ?? 0,
    (s.providerSessionIds ?? []).join("; "),
    s.provider ?? "",
    s.reasoningTokens ?? "", s.reasoningReportedTurns ?? 0,
    s.provider && s.providerSessionIds?.length === 1 && s.sessionId === providerSessionKey(s.provider, s.providerSessionIds[0]!) ? "CLI session" : "legacy chat",
  ]);
  return `\uFEFF${rows.map((row) => row.map(cell).join(",")).join("\r\n")}\r\n`;
}
