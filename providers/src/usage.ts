import type { TokenUsage } from "../../protocol/src/index.js";

/** Informational /usage prints rounded credit totals, distinct from legacy premium requests. */
export function parseCopilotCredits(report: string): number | null {
  const clean = report.replace(/\u001b\[[0-9;]*m/g, "");
  const match = clean.match(/(?:^|\n)\s*Requests:\s*(\d[\d,]*(?:\.\d+)?)\s+AI Credits\b/i);
  if (!match) return null;
  const total = Number(match[1]!.replace(/,/g, ""));
  return Number.isFinite(total) && total >= 0 ? total : null;
}

/** One counter per CLI session: ACP reports session totals at prompt completion. */
export class CumulativeUsage {
  private previous: TokenUsage | null = null;

  next(total: TokenUsage): TokenUsage {
    const before = this.previous;
    this.previous = { ...total };
    if (!before || total.totalTokens < before.totalTokens) return { ...total };
    return {
      inputTokens: Math.max(0, total.inputTokens - before.inputTokens),
      outputTokens: Math.max(0, total.outputTokens - before.outputTokens),
      cachedInputTokens: Math.max(0, total.cachedInputTokens - before.cachedInputTokens),
      totalTokens: Math.max(0, total.totalTokens - before.totalTokens),
    };
  }
}
