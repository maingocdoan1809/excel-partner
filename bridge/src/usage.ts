import { appendFileSync, mkdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { z } from "zod";
import { TokenUsageSchema, type UsageSummary } from "../../protocol/src/index.js";

const RecordSchema = z.object({
  sessionId: z.string(),
  requestId: z.string(),
  provider: z.string(),
  model: z.string().nullable(),
  startedAt: z.number(),
  durationMs: z.number().nonnegative(),
  reason: z.enum(["completed", "cancelled", "error"]),
  usage: TokenUsageSchema.nullable(),
});
export type UsageRecord = z.infer<typeof RecordSchema>;

/** Append-only local ledger. No prompts, workbook data or credentials. */
export class UsageStore {
  private readonly records: UsageRecord[] = [];

  constructor(private readonly file: string | null) {
    if (!file) return;
    let contents: string;
    try { contents = readFileSync(file, "utf8"); }
    catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
      throw error;
    }
    for (const line of contents.split("\n")) {
      if (!line.trim()) continue;
      try {
        const parsed = RecordSchema.safeParse(JSON.parse(line));
        if (parsed.success) this.records.push(parsed.data);
      } catch { /* A damaged line must not discard the rest of the ledger. */ }
    }
  }

  append(record: UsageRecord): void {
    if (this.file) {
      mkdirSync(path.dirname(this.file), { recursive: true, mode: 0o700 });
      // A leading newline also isolates an incomplete final line after a crash.
      appendFileSync(this.file, `\n${JSON.stringify(record)}\n`, { mode: 0o600 });
    }
    this.records.push(record);
  }

  summaries(): UsageSummary[] {
    const sessions = new Map<string, UsageSummary>();
    for (const record of this.records) {
      let summary = sessions.get(record.sessionId);
      if (!summary) {
        summary = { sessionId: record.sessionId, turns: 0, reportedTurns: 0, creditReportedTurns: 0, durationMs: 0,
          inputTokens: 0, outputTokens: 0, cachedInputTokens: 0, totalTokens: 0,
          startedAt: record.startedAt, updatedAt: record.startedAt };
        sessions.set(record.sessionId, summary);
      }
      summary.turns++;
      summary.durationMs += record.durationMs;
      summary.startedAt = Math.min(summary.startedAt, record.startedAt);
      summary.updatedAt = Math.max(summary.updatedAt, record.startedAt + record.durationMs);
      if (record.usage) {
        summary.reportedTurns++;
        summary.inputTokens += record.usage.inputTokens;
        summary.outputTokens += record.usage.outputTokens;
        summary.cachedInputTokens += record.usage.cachedInputTokens;
        summary.totalTokens += record.usage.totalTokens;
        if (record.usage.copilotCredits !== undefined) {
          summary.creditReportedTurns++;
          summary.copilotCredits = (summary.copilotCredits ?? 0) + record.usage.copilotCredits;
        }
      }
    }
    return [...sessions.values()].sort((a, b) => b.updatedAt - a.updatedAt);
  }
}
