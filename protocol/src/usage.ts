import { z } from "zod";

/** Input includes cached tokens; cachedInputTokens is a subset, not added again. */
export const TokenUsageSchema = z.object({
  inputTokens: z.number().int().nonnegative(),
  outputTokens: z.number().int().nonnegative(),
  cachedInputTokens: z.number().int().nonnegative(),
  totalTokens: z.number().int().nonnegative(),
  /** Optional cache split. Both cache buckets are already included in inputTokens. */
  cacheWriteTokens: z.number().int().nonnegative().optional(),
  uncachedInputTokens: z.number().int().nonnegative().optional(),
  /** Original integer from Copilot session.usage_record; 1 credit = 1e9 nano AI units. */
  copilotCreditNanoAiu: z.number().int().nonnegative().safe().optional(),
  source: z.enum(["copilot_events", "copilot_acp", "claude_result", "codex_thread", "opencode_export"]).optional(),
  /** Subset of outputTokens after provider-specific normalisation. */
  reasoningTokens: z.number().int().nonnegative().optional(),
  providerSessionId: z.string().optional(),
  /** Legacy records can contain rounded /usage credits. New records retain nano units. */
  copilotCredits: z.number().nonnegative().optional(),
});
export type TokenUsage = z.infer<typeof TokenUsageSchema>;

/** Provider namespace prevents two CLIs with the same native ID sharing accounting. */
export function providerSessionKey(provider: string, nativeId: string): string { return `${provider}:${nativeId}`; }

export interface UsageSummary extends TokenUsage {
  sessionId: string;
  provider?: string;
  title?: string;
  turns: number;
  reportedTurns: number;
  creditReportedTurns: number;
  exactCreditReportedTurns?: number;
  cacheSplitReportedTurns?: number;
  originalReportedTurns?: number;
  reasoningReportedTurns?: number;
  providerSessionIds?: string[];
  durationMs: number;
  startedAt: number;
  updatedAt: number;
}
