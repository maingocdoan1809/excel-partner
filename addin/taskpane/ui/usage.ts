import type { UsageSummary } from "../../../protocol/src/index.js";

export function formatDuration(ms: number): string {
  const seconds = Math.floor(Math.max(0, ms) / 1000);
  return seconds >= 3600
    ? `${Math.floor(seconds / 3600)}h ${Math.floor(seconds % 3600 / 60)}m ${seconds % 60}s`
    : `${Math.floor(seconds / 60)}m ${seconds % 60}s`;
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
    "Copilot AI credits (approx, CLI /usage)", "Turns with credit data", "Credit data status",
  ]];
  for (const s of summaries) rows.push([
    s.sessionId, titles.get(s.sessionId) ?? "", new Date(s.startedAt).toISOString(), new Date(s.updatedAt).toISOString(),
    s.turns, s.reportedTurns, (s.durationMs / 1000).toFixed(3), (s.durationMs / Math.max(1, s.turns) / 1000).toFixed(3),
    s.reportedTurns ? s.inputTokens : "", s.reportedTurns ? s.outputTokens : "",
    s.reportedTurns ? s.cachedInputTokens : "", s.reportedTurns ? s.totalTokens : "",
    s.reportedTurns === s.turns ? "complete" : s.reportedTurns ? "partial" : "unavailable",
    s.creditReportedTurns ? (s.copilotCredits ?? 0).toFixed(2) : "", s.creditReportedTurns ?? 0,
    s.creditReportedTurns === s.turns ? "complete (rounded)" : s.creditReportedTurns ? "partial (rounded)" : "unavailable",
  ]);
  return `\uFEFF${rows.map((row) => row.map(cell).join(",")).join("\r\n")}\r\n`;
}
