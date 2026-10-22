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
      ...(total.source ? { source: total.source } : {}),
      ...(total.providerSessionId ? { providerSessionId: total.providerSessionId } : {}),
      inputTokens: Math.max(0, total.inputTokens - before.inputTokens),
      outputTokens: Math.max(0, total.outputTokens - before.outputTokens),
      cachedInputTokens: Math.max(0, total.cachedInputTokens - before.cachedInputTokens),
      totalTokens: Math.max(0, total.totalTokens - before.totalTokens),
      ...optionalDelta(total, before, "cacheWriteTokens"),
      ...optionalDelta(total, before, "uncachedInputTokens"),
      ...optionalDelta(total, before, "copilotCreditNanoAiu"),
      ...optionalDelta(total, before, "reasoningTokens"),
      ...(total.copilotCreditNanoAiu !== undefined && before.copilotCreditNanoAiu !== undefined
        ? { copilotCredits: Math.max(0, total.copilotCreditNanoAiu - before.copilotCreditNanoAiu) / 1e9 } : {}),
    };
  }
}

function optionalDelta(total: TokenUsage, before: TokenUsage,
  key: "cacheWriteTokens" | "uncachedInputTokens" | "copilotCreditNanoAiu" | "reasoningTokens"): Partial<TokenUsage> {
  return total[key] !== undefined && before[key] !== undefined
    ? { [key]: Math.max(0, total[key]! - before[key]!) } : {};
}
