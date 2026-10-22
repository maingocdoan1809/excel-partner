import { z } from "zod";
import { TokenUsageSchema, type TokenUsage } from "../../protocol/src/index.js";
import { nativeTitle } from "./session-metadata.js";

const Count = z.number().int().nonnegative().safe();
const Tokens = z.object({ input: Count, output: Count, reasoning: Count,
  cache: z.object({ read: Count, write: Count }) });

/** OpenCode input excludes both cache buckets, and output excludes reasoning.
 * Its ACP prompt response describes only the final model call, not the full user turn.
 * Use export's session totals (or every assistant message on older releases).
 */
export function parseOpencodeExport(data: unknown, id: string): { usage: TokenUsage; title?: string } {
  const parsed = z.object({ info: z.object({ id: z.string(), title: z.string().optional(), tokens: Tokens.optional() }),
    messages: z.array(z.object({ info: z.object({ id: z.string(), role: z.string(), tokens: Tokens.optional() }) })).optional() }).parse(data);
  if (parsed.info.id !== id) throw new Error("OpenCode exported a different session");
  let totals = parsed.info.tokens;
  if (!totals) {
    totals = { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } };
    const seen = new Set<string>();
    let found = false;
    for (const { info } of parsed.messages ?? []) {
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
