import fs from "node:fs/promises";
import { createReadStream } from "node:fs";
import { createInterface } from "node:readline";
import path from "node:path";
import os from "node:os";
import { parse } from "yaml";
import type { TokenUsage } from "../../protocol/src/index.js";

export function nativeTitle(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  return value.replace(/[\x00-\x1f\x7f]/g, " ").replace(/\s+/g, " ").trim().slice(0, 200) || undefined;
}

export function validNativeId(id: string): boolean { return /^[a-zA-Z0-9_-]{1,128}$/.test(id); }

export async function copilotTitle(id: string, env: NodeJS.ProcessEnv): Promise<string | undefined> {
  if (!validNativeId(id)) return undefined;
  const file = path.join(env.COPILOT_HOME ?? path.join(os.homedir(), ".copilot"), "session-state", id, "workspace.yaml");
  try {
    if ((await fs.stat(file)).size > 64 * 1024) return undefined;
    return nativeTitle(parse(await fs.readFile(file, "utf8"), { maxAliasCount: 0 })?.name);
  } catch { return undefined; }
}

export async function claudeTitle(id: string, cwd: string, env: NodeJS.ProcessEnv): Promise<string | undefined> {
  if (!validNativeId(id)) return undefined;
  const file = path.join(env.CLAUDE_CONFIG_DIR ?? path.join(os.homedir(), ".claude"), "projects",
    cwd.replace(/[^a-zA-Z0-9]/g, "-"), `${id}.jsonl`);
  let title: string | undefined;
  let customTitle: string | undefined;
  try {
    for await (const line of createInterface({ input: createReadStream(file), crlfDelay: Infinity })) {
      try {
        const event = JSON.parse(line);
        if (event.sessionId && event.sessionId !== id) continue;
        if (event.type === "custom-title") { customTitle = nativeTitle(event.customTitle) ?? customTitle; continue; }
        const value = event.type === "ai-title" ? event.aiTitle : event.type === "summary" ? event.summary : undefined;
        title = nativeTitle(value) ?? title;
      } catch { /* An incomplete last line does not hide an existing native title. */ }
    }
  } catch { return undefined; }
  return customTitle ?? title;
}

/** Codex rollout token_count is the cumulative odometer, including work before resume. */
export async function codexTokenBaseline(file: string | null | undefined): Promise<TokenUsage | undefined> {
  if (!file) return undefined;
  let total: TokenUsage | undefined;
  try {
    for await (const line of createInterface({ input: createReadStream(file), crlfDelay: Infinity })) {
      try {
        const event = JSON.parse(line);
        const u = event.type === "token_usage_record" ? event.payload?.thread_token_usage
          : event.type === "event_msg" && event.payload?.type === "token_count" ? event.payload.info?.total_token_usage : undefined;
        if (!u || ![u.input_tokens, u.output_tokens, u.total_tokens].every(n => Number.isSafeInteger(n) && n >= 0)) continue;
        total = { inputTokens: u.input_tokens, outputTokens: u.output_tokens, totalTokens: u.total_tokens,
          cachedInputTokens: u.cached_input_tokens ?? 0,
          ...(Number.isSafeInteger(u.cache_write_input_tokens) && u.cache_write_input_tokens >= 0 && (u.cached_input_tokens ?? 0) + u.cache_write_input_tokens <= u.input_tokens
            ? { cacheWriteTokens: u.cache_write_input_tokens, uncachedInputTokens: u.input_tokens - (u.cached_input_tokens ?? 0) - u.cache_write_input_tokens } : {}),
          ...(Number.isSafeInteger(u.reasoning_output_tokens) && u.reasoning_output_tokens >= 0 ? { reasoningTokens: u.reasoning_output_tokens } : {}) };
      } catch { /* Ignore a partially written final line. */ }
    }
  } catch { return undefined; }
  return total;
}
