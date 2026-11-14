import { appendFileSync, mkdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import { z } from "zod";
import { TokenUsageSchema, type TokenUsage, type UsageSummary } from "../../protocol/src/index.js";

const RecordSchema = z.object({
  sessionId: z.string(),
  requestId: z.string(),
  provider: z.string(),
  model: z.string().nullable(),
  providerSessionId: z.string().optional(),
  title: z.string().max(200).optional(),
  startedAt: z.number(),
  durationMs: z.number().nonnegative(),
  reason: z.enum(["completed", "cancelled", "error"]),
  usage: TokenUsageSchema.nullable(),
});
export type UsageRecord = z.infer<typeof RecordSchema>;

/**
 * What the provider's CLI reported for a whole conversation at one moment. For a CLI whose usage is read when someone looks at
 * it (see `ProviderAdapter.sessionUsage`) the turns carry none; this stands for them, up to `at`.
 */
const TotalSchema = z.object({
  total: z.literal(true),
  provider: z.string(),
  providerSessionId: z.string(),
  at: z.number(),
  usage: TokenUsageSchema,
});
type Total = z.infer<typeof TotalSchema>;

const COUNTS = ["inputTokens", "outputTokens", "cachedInputTokens", "totalTokens", "cacheWriteTokens", "uncachedInputTokens", "reasoningTokens"] as const;

/** The ledger file is rewritten with only what is kept in memory once it grows past this; the previous file stays as `.1`. */
export const USAGE_FILE_MAX_BYTES = 5 * 1024 * 1024;
/** The usage panel shows recent chats: turns this much older than the newest, or beyond the newest `USAGE_MAX_RECORDS`, are dropped. */
export const USAGE_MAX_AGE_MS = 90 * 24 * 60 * 60_000;
export const USAGE_MAX_RECORDS = 5_000;

/** Append-only local ledger, bounded in memory and on disk. No prompts, workbook data or credentials. */
export class UsageStore {
  private records: UsageRecord[] = [];
  private readonly totals = new Map<string, Total>();

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
        const json = JSON.parse(line);
        const total = TotalSchema.safeParse(json);
        if (total.success) { this.keepTotal(total.data); continue; }
        const parsed = RecordSchema.safeParse(json);
        if (parsed.success) this.records.push(parsed.data);
      } catch { /* A damaged line must not discard the rest of the ledger. */ }
    }
    this.trim();
    this.rotateIfLarge();
  }

  append(record: UsageRecord): void {
    this.write(record);
    this.records.push(record);
    this.trim();
    this.rotateIfLarge();
  }

  /** Measured from the newest turn, so a ledger that was not used for a while still shows its last chats. */
  private trim(): void {
    const oldest = (this.records.at(-1)?.startedAt ?? 0) - USAGE_MAX_AGE_MS;
    if (this.records.length > USAGE_MAX_RECORDS || (this.records[0] && this.records[0].startedAt < oldest)) {
      this.records = this.records.filter((record) => record.startedAt >= oldest).slice(-USAGE_MAX_RECORDS);
    }
  }

  /** Keep the previous file as `.1` and start a new one holding only what is kept in memory. */
  private rotateIfLarge(): void {
    if (!this.file) return;
    try {
      if (statSync(this.file).size <= USAGE_FILE_MAX_BYTES) return;
      rmSync(`${this.file}.1`, { force: true });
      renameSync(this.file, `${this.file}.1`);
      const lines = [...this.totals.values(), ...this.records].map((entry) => JSON.stringify(entry));
      writeFileSync(this.file, lines.join("\n") + "\n", { mode: 0o600 });
    } catch { /* The ledger keeps growing until rotation succeeds; it is never lost. */ }
  }

  private write(entry: object): void {
    if (!this.file) return;
    mkdirSync(path.dirname(this.file), { recursive: true, mode: 0o700 });
    // A leading newline also isolates an incomplete final line after a crash.
    appendFileSync(this.file, `\n${JSON.stringify(entry)}\n`, { mode: 0o600 });
  }

  private keepTotal(total: Total): void {
    const key = `${total.provider}:${total.providerSessionId}`;
    if ((this.totals.get(key)?.at ?? 0) <= total.at) this.totals.set(key, total);
  }

  /** Record the CLI's own total for a conversation, read just now. */
  settle(provider: string, providerSessionId: string, usage: TokenUsage, at = Date.now()): void {
    const total: Total = { total: true, provider, providerSessionId, at, usage };
    this.write(total);
    this.keepTotal(total);
  }

  /** The conversations (of `sessionId`, or all) with turns the latest recorded total does not cover yet. */
  unsettled(sessionId?: string | null): Array<{ provider: string; providerSessionId: string }> {
    const found = new Map<string, { provider: string; providerSessionId: string }>();
    for (const record of this.records) {
      if (record.usage || !record.providerSessionId || (sessionId && record.sessionId !== sessionId)) continue;
      const key = `${record.provider}:${record.providerSessionId}`;
      if ((this.totals.get(key)?.at ?? 0) < record.startedAt + record.durationMs) found.set(key, { provider: record.provider, providerSessionId: record.providerSessionId });
    }
    return [...found.values()];
  }

  summaries(): UsageSummary[] {
    const sessions = new Map<string, UsageSummary>();
    const legacyCredits = new Map<string, number>();
    /** Per chat and CLI conversation: what the turns reported themselves, and how many turns reported nothing but are covered by a recorded total. */
    const settled = new Map<string, Map<string, { provider: string; recorded: Partial<Record<typeof COUNTS[number], number>>; covered: number }>>();
    for (const record of this.records) {
      let summary = sessions.get(record.sessionId);
      if (!summary) {
        summary = { sessionId: record.sessionId, turns: 0, reportedTurns: 0, creditReportedTurns: 0, durationMs: 0,
          provider: record.provider,
          exactCreditReportedTurns: 0, cacheSplitReportedTurns: 0, originalReportedTurns: 0, reasoningReportedTurns: 0, providerSessionIds: [],
          inputTokens: 0, outputTokens: 0, cachedInputTokens: 0, totalTokens: 0,
          startedAt: record.startedAt, updatedAt: record.startedAt };
        sessions.set(record.sessionId, summary);
      }
      summary.turns++;
      if (record.providerSessionId) {
        const natives = (settled.get(record.sessionId) ?? settled.set(record.sessionId, new Map()).get(record.sessionId)!);
        const entry = natives.get(record.providerSessionId) ?? { provider: record.provider, recorded: {} as Partial<Record<typeof COUNTS[number], number>>, covered: 0 };
        natives.set(record.providerSessionId, entry);
        if (record.usage) for (const key of COUNTS) entry.recorded[key] = (entry.recorded[key] ?? 0) + (record.usage[key] ?? 0);
        else if ((this.totals.get(`${record.provider}:${record.providerSessionId}`)?.at ?? 0) >= record.startedAt + record.durationMs) entry.covered++;
      }
      if (record.title) summary.title = record.title;
      if (record.providerSessionId && !summary.providerSessionIds!.includes(record.providerSessionId)) summary.providerSessionIds!.push(record.providerSessionId);
      summary.durationMs += record.durationMs;
      summary.startedAt = Math.min(summary.startedAt, record.startedAt);
      summary.updatedAt = Math.max(summary.updatedAt, record.startedAt + record.durationMs);
      if (record.usage) {
        summary.reportedTurns++;
        summary.inputTokens += record.usage.inputTokens;
        summary.outputTokens += record.usage.outputTokens;
        summary.cachedInputTokens += record.usage.cachedInputTokens;
        summary.totalTokens += record.usage.totalTokens;
        if (record.usage.reasoningTokens !== undefined) {
          summary.reasoningReportedTurns!++;
          summary.reasoningTokens = (summary.reasoningTokens ?? 0) + record.usage.reasoningTokens;
        }
        if (record.usage.source && record.usage.source !== "copilot_acp") summary.originalReportedTurns!++;
        const providerId = record.usage.providerSessionId;
        if (providerId && !summary.providerSessionIds!.includes(providerId)) summary.providerSessionIds!.push(providerId);
        if (record.usage.cacheWriteTokens !== undefined && record.usage.uncachedInputTokens !== undefined) {
          summary.cacheSplitReportedTurns!++;
          summary.cacheWriteTokens = (summary.cacheWriteTokens ?? 0) + record.usage.cacheWriteTokens;
          summary.uncachedInputTokens = (summary.uncachedInputTokens ?? 0) + record.usage.uncachedInputTokens;
        }
        if (record.usage.copilotCredits !== undefined) {
          summary.creditReportedTurns++;
          if (record.usage.copilotCreditNanoAiu !== undefined) {
            summary.exactCreditReportedTurns!++;
            summary.copilotCreditNanoAiu = (summary.copilotCreditNanoAiu ?? 0) + record.usage.copilotCreditNanoAiu;
          } else {
            legacyCredits.set(record.sessionId, (legacyCredits.get(record.sessionId) ?? 0) + record.usage.copilotCredits);
          }
          summary.copilotCredits = (summary.copilotCreditNanoAiu ?? 0) / 1e9 + (legacyCredits.get(record.sessionId) ?? 0);
        }
      }
    }
    // Turns that reported nothing are accounted for by the CLI's own total for their conversation, less what other turns of it reported.
    for (const [sessionId, natives] of settled) {
      const summary = sessions.get(sessionId)!;
      for (const [nativeId, entry] of natives) {
        const total = this.totals.get(`${entry.provider}:${nativeId}`)?.usage;
        if (!entry.covered || !total) continue;
        for (const key of COUNTS) {
          if (total[key] === undefined) continue;
          const extra = Math.max(0, total[key]! - (entry.recorded[key] ?? 0));
          summary[key] = (summary[key] ?? 0) + extra;
        }
        summary.reportedTurns += entry.covered;
        summary.originalReportedTurns = (summary.originalReportedTurns ?? 0) + entry.covered;
        if (total.reasoningTokens !== undefined) summary.reasoningReportedTurns = (summary.reasoningReportedTurns ?? 0) + entry.covered;
        if (total.cacheWriteTokens !== undefined && total.uncachedInputTokens !== undefined) summary.cacheSplitReportedTurns = (summary.cacheSplitReportedTurns ?? 0) + entry.covered;
      }
    }
    return [...sessions.values()].sort((a, b) => b.updatedAt - a.updatedAt);
  }
}
