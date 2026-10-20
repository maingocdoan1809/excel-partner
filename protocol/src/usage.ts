import { z } from "zod";

/** Input includes cached tokens; cachedInputTokens is a subset, not added again. */
export const TokenUsageSchema = z.object({
  inputTokens: z.number().int().nonnegative(),
  outputTokens: z.number().int().nonnegative(),
  cachedInputTokens: z.number().int().nonnegative(),
  totalTokens: z.number().int().nonnegative(),
  /** Per-turn increment from Copilot /usage; rounded by the CLI, not an invoice. */
  copilotCredits: z.number().nonnegative().optional(),
});
export type TokenUsage = z.infer<typeof TokenUsageSchema>;

export interface UsageSummary extends TokenUsage {
  sessionId: string;
  turns: number;
  reportedTurns: number;
  creditReportedTurns: number;
  durationMs: number;
  startedAt: number;
  updatedAt: number;
}
