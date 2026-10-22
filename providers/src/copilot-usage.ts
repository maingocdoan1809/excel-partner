import { open } from "node:fs/promises";
import { z } from "zod";
import { TokenUsageSchema, type TokenUsage } from "../../protocol/src/index.js";

const Count = z.number().int().nonnegative().safe();
const RecordUsage = z.object({
  inputTokens: Count,
  outputTokens: Count,
  cacheReadTokens: Count.optional(),
  cacheWriteTokens: Count.optional(),
  reasoningTokens: Count.optional(),
  copilotUsage: z.object({ totalNanoAiu: Count }).optional(),
  aiCreditsStatus: z.string().optional(),
});

/** Extract only accounting from the exact CLI session; never persist prompt content.
 * Checkpoints/shutdown contain duplicates of the same accounting and are not added.
 * Retain a cumulative snapshot so switching from ACP cannot count tokens twice.
 */
export class CopilotUsageReader {
  private offset = 0;
  private readonly seen = new Set<string>();
  private input = 0;
  private output = 0;
  private cacheRead = 0;
  private cacheWrite = 0;
  private nano = 0;
  private records = 0;
  private cacheComplete = true;
  private reasoningComplete = true;
  private reasoning = 0;
  private creditComplete = true;

  constructor(private readonly file: string, private readonly sessionId: string) {}

  async snapshot(): Promise<TokenUsage | null> {
    let handle;
    try { handle = await open(this.file, "r"); }
    catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
      throw error;
    }
    try {
      const size = (await handle.stat()).size;
      if (size < this.offset) throw new Error("Copilot usage file was truncated");
      const chunk = Buffer.alloc(64 * 1024);
      let position = this.offset;
      let pending = Buffer.alloc(0);
      while (position < size) {
        const { bytesRead } = await handle.read(chunk, 0, Math.min(chunk.length, size - position), position);
        if (!bytesRead) break;
        position += bytesRead;
        pending = Buffer.concat([pending, chunk.subarray(0, bytesRead)]);
        let start = 0;
        let end: number;
        while ((end = pending.indexOf(10, start)) !== -1) {
          const line = pending.subarray(start, end).toString("utf8");
          this.consume(line);
          this.offset += end + 1 - start;
          start = end + 1;
        }
        pending = pending.subarray(start);
        if (pending.length > 8 * 1024 * 1024) throw new Error("Copilot event line exceeds 8 MB");
      }
      // A partial final line stays unread until the CLI finishes writing it.
    } finally { await handle.close(); }
    if (!this.records) return null;
    return TokenUsageSchema.parse({
      inputTokens: this.input, outputTokens: this.output,
      cachedInputTokens: this.cacheRead, totalTokens: this.input + this.output,
      ...(this.cacheComplete ? { cacheWriteTokens: this.cacheWrite,
        uncachedInputTokens: this.input - this.cacheRead - this.cacheWrite } : {}),
      ...(this.creditComplete ? { copilotCreditNanoAiu: this.nano, copilotCredits: this.nano / 1e9 } : {}),
      source: "copilot_events", providerSessionId: this.sessionId,
      ...(this.reasoningComplete ? { reasoningTokens: this.reasoning } : {}),
    });
  }

  private consume(line: string): void {
    if (!line.trim()) return;
    const event = JSON.parse(line);
    if (event.type !== "session.usage_record") return;
    const usage = RecordUsage.parse(event.data?.usage);
    const accounting = event.data?.usage?.accounting;
    // Replayed usage must not be counted again, nor attributed from another session.
    if (accounting?.sourceSessionId && accounting.sourceSessionId !== this.sessionId) return;
    const id = accounting?.usageId ?? event.id;
    if (typeof id !== "string") throw new Error("Copilot usage record has no identity");
    if (this.seen.has(id)) return;
    const read = usage.cacheReadTokens ?? 0;
    const write = usage.cacheWriteTokens ?? 0;
    if (read + write > usage.inputTokens) throw new Error("Copilot cache exceeds input tokens");
    const nano = usage.copilotUsage?.totalNanoAiu;
    this.seen.add(id);
    this.records++;
    this.input += usage.inputTokens;
    this.output += usage.outputTokens;
    this.cacheRead += read;
    this.cacheWrite += write;
    this.nano += nano ?? 0;
    this.reasoning += usage.reasoningTokens ?? 0;
    this.reasoningComplete &&= usage.reasoningTokens !== undefined;
    this.cacheComplete &&= usage.cacheReadTokens !== undefined && usage.cacheWriteTokens !== undefined;
    this.creditComplete &&= nano !== undefined && usage.aiCreditsStatus === "complete";
  }
}
