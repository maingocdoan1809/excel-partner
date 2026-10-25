import { z } from "zod";
import { TokenUsageSchema, type TokenUsage } from "../../protocol/src/index.js";
import { nativeTitle } from "./session-metadata.js";

const Count = z.number().int().nonnegative().safe();
const Tokens = z.object({ input: Count, output: Count, reasoning: Count,
  cache: z.object({ read: Count, write: Count }) });

const LegacyMessage = z.object({ info: z.object({ id: z.string(), role: z.string(), tokens: Tokens.optional() }) });
const V2Message = z.object({ id: z.string(), type: z.string(), tokens: Tokens.optional() });
const ExportInfo = z.object({ id: z.string(), title: z.string().optional(), tokens: Tokens.optional() });
const LegacyExport = z.object({ info: ExportInfo, messages: z.array(LegacyMessage).optional() });
const V2Export = z.object({ info: ExportInfo, messages: z.array(z.unknown()).optional() });

/** OpenCode input excludes both cache buckets, and output excludes reasoning.
 * Its ACP prompt response describes only the final model call, not the full user turn.
 * Use export's session totals (or every assistant message on older releases).
 */
export function parseOpencodeExport(data: unknown, id: string, v2 = false): { usage: TokenUsage; title?: string } {
  const parsed = v2 ? V2Export.parse(data) : LegacyExport.parse(data);
  if (parsed.info.id !== id) throw new Error("OpenCode exported a different session");
  let totals = parsed.info.tokens;
  if (!totals) {
    totals = { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } };
    const seen = new Set<string>();
    let found = false;
    for (const message of parsed.messages ?? []) {
      // V1 nests assistant data under info; V2 stores typed messages directly.
      const native = v2 ? V2Message.parse(message) : undefined;
      const info = native ? { ...native, role: native.type } : LegacyMessage.parse(message).info;
      if (info.role !== "assistant" || seen.has(info.id)) continue;
      seen.add(info.id);
      if (!info.tokens) throw new Error("OpenCode assistant message is missing original usage");
      found = true;
      totals.input += info.tokens.input;
      totals.output += info.tokens.output;
      totals.reasoning += info.tokens.reasoning;
      totals.cache.read += info.tokens.cache.read;
      totals.cache.write += info.tokens.cache.write;
    }
    if (!found) throw new Error("OpenCode original session usage is not available yet");
  }
  const inputTokens = totals.input + totals.cache.read + totals.cache.write;
  const outputTokens = totals.output + totals.reasoning;
  return { title: nativeTitle(parsed.info.title), usage: TokenUsageSchema.parse({
    inputTokens, outputTokens, totalTokens: inputTokens + outputTokens,
    cachedInputTokens: totals.cache.read, cacheWriteTokens: totals.cache.write,
    uncachedInputTokens: totals.input, reasoningTokens: totals.reasoning,
    source: "opencode_export", providerSessionId: id,
  }) };
}
