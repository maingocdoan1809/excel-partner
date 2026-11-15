import fs from "node:fs/promises";
import path from "node:path";
import type { ExternalSessionSummary, ImportedSession } from "../../../protocol/src/external-sessions.js";
import { MAX_LISTED_SESSIONS, TranscriptBuilder, isOwnWorkspace, mtimeOf, newestFirst, samePath, type SessionHomes, scanJsonLines, titleFromText, toMillis, userHome } from "../external-sessions.js";
import { nativeTitle, validNativeId } from "../session-metadata.js";

const HEAD_BYTES = 1024 * 1024;
/** Codex writes `source` on the first line of a rollout: where the conversation was held. */
const ORIGINS: Record<string, string> = { cli: "Terminal", exec: "Terminal", vscode: "VS Code" };
/** What Codex puts at the top of a conversation as if the person had said it. */
const INJECTED = /^<(?:environment_context|user_instructions|user_action|turn_aborted)\b|^# AGENTS\.md instructions/;

export const codexHome = (env: NodeJS.ProcessEnv): string => env.CODEX_HOME ?? path.join(userHome(env), ".codex");

interface Rollout { file: string; mtime: number }

/** Rollouts live in sessions/YYYY/MM/DD/rollout-<time>-<id>.jsonl. */
async function rollouts(env: NodeJS.ProcessEnv): Promise<Rollout[]> {
  const found: Rollout[] = [];
  const walk = async (dir: string, depth: number): Promise<void> => {
    let entries: import("node:fs").Dirent[] = [];
    try { entries = await fs.readdir(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory() && depth < 4) await walk(full, depth + 1);
      else if (entry.isFile() && /^rollout-.*\.jsonl$/.test(entry.name)) { const mtime = await mtimeOf(full); if (mtime !== undefined) found.push({ file: full, mtime }); }
    }
  };
  await walk(path.join(codexHome(env), "sessions"), 0);
  return found;
}

/** Names the CLI gave conversations (`/rename`, automatic titles), by thread id. */
async function threadNames(env: NodeJS.ProcessEnv): Promise<Map<string, string>> {
  const names = new Map<string, string>();
  try { await scanJsonLines(path.join(codexHome(env), "session_index.jsonl"), (line) => { const name = nativeTitle(line.thread_name); if (typeof line.id === "string" && name) names.set(line.id, name); }); } catch { /* No index: titles come from the first prompt. */ }
  return names;
}

const injected = (text: string): boolean => INJECTED.test(text.trimStart());

/** The text of a `response_item` message, whichever of the content kinds Codex wrote it in. */
const messageText = (payload: any): string => Array.isArray(payload.content)
  ? payload.content.map((part: any) => typeof part?.text === "string" ? part.text : "").join("") : typeof payload.content === "string" ? payload.content : "";

async function summarize(item: Rollout, names: Map<string, string>): Promise<ExternalSessionSummary | null> {
  let id: string | undefined, cwd: string | undefined, source: unknown, prompt = "", fallback = "";
  await scanJsonLines(item.file, (line) => {
    const payload = line.payload;
    if (line.type === "session_meta" && payload) { id = typeof payload.id === "string" ? payload.id : id; cwd = typeof payload.cwd === "string" ? payload.cwd : cwd; source = payload.source; }
    else if (line.type === "event_msg" && payload?.type === "user_message" && typeof payload.message === "string") { prompt = prompt || payload.message; }
    else if (line.type === "response_item" && payload?.type === "message" && payload.role === "user") { const text = messageText(payload); if (!fallback && text && !injected(text)) fallback = text; }
    return !(id && (prompt || fallback));
  }, { maxBytes: HEAD_BYTES });
  const first = prompt || fallback;
  // A conversation a sub-agent held (its source is an object, not a name) is part of its parent's.
  if (!id || !validNativeId(id) || !first || (source !== undefined && typeof source !== "string") || isOwnWorkspace(cwd)) return null;
  return { id, title: names.get(id) ?? titleFromText(first), origin: ORIGINS[String(source)] ?? "Codex", ...(cwd ? { cwd } : {}), updatedAt: item.mtime };
}

export async function listCodexSessions(env: NodeJS.ProcessEnv): Promise<ExternalSessionSummary[]> {
  const [files, names] = await Promise.all([rollouts(env), threadNames(env)]);
  const summaries = await Promise.all(newestFirst(files, MAX_LISTED_SESSIONS).map((item) => summarize(item, names).catch(() => null)));
  return summaries.filter((item): item is ExternalSessionSummary => item !== null);
}

/** What a tool call returned: Codex sometimes wraps the text as `{"output": "…"}`. */
function toolOutput(value: unknown): string | undefined {
  const text = typeof value === "string" ? value : value && typeof (value as any).content === "string" ? (value as any).content : undefined;
  if (text === undefined) return undefined;
  try { const parsed = JSON.parse(text); if (parsed && typeof parsed.output === "string") return parsed.output; } catch { /* Plain text. */ }
  return text;
}

/** Codex resumes a thread from its own `sessions` folder: a private state folder gets a copy of the rollout, at the same place. */
async function copyRollout(homes: SessionHomes, file: string): Promise<void> {
  const from = path.join(codexHome(homes.source), "sessions"), to = path.join(codexHome(homes.target), "sessions");
  if (samePath(from, to)) return;
  const destination = path.join(to, path.relative(from, file));
  await fs.mkdir(path.dirname(destination), { recursive: true });
  try { await fs.copyFile(file, destination, fs.constants.COPYFILE_EXCL); }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error; }
}

/** Reads a Codex conversation; the CLI itself resumes it by thread id, so only a private state folder needs the rollout copied in. */
export async function importCodexSession(homes: SessionHomes, id: string): Promise<ImportedSession> {
  if (!validNativeId(id)) throw new Error("Invalid Codex thread ID");
  const names = await threadNames(homes.source);
  const match = (await rollouts(homes.source)).find((item) => item.file.endsWith(`-${id}.jsonl`));
  if (!match) throw new Error("That Codex conversation was not found on this computer.");
  // Newer rollouts record each message twice, as an event and as a response item: the events are the clean text.
  let hasEvents = false;
  await scanJsonLines(match.file, (line) => { hasEvents = line.type === "event_msg" && line.payload?.type === "user_message"; return !hasEvents; }, { maxBytes: HEAD_BYTES });
  const transcript = new TranscriptBuilder();
  let startedAt: number | undefined, updatedAt: number | undefined;
  await scanJsonLines(match.file, (line) => {
    const at = toMillis(line.timestamp);
    if (at !== undefined) { startedAt ??= at; updatedAt = at; }
    const payload = line.payload;
    if (!payload) return;
    if (line.type === "event_msg" && hasEvents) {
      if (payload.type === "user_message") transcript.user(payload.message);
      else if (payload.type === "agent_message") transcript.assistant(payload.message);
      return;
    }
    if (line.type !== "response_item") return;
    if (payload.type === "message" && !hasEvents) {
      const text = messageText(payload);
      if (payload.role === "user" && !injected(text)) transcript.user(text);
      else if (payload.role === "assistant") transcript.assistant(text);
    } else if (payload.type === "function_call") transcript.toolCall(String(payload.call_id), String(payload.name ?? "tool"), payload.arguments);
    else if (payload.type === "custom_tool_call") transcript.toolCall(String(payload.call_id), String(payload.name ?? "tool"), payload.input);
    else if (payload.type === "local_shell_call") transcript.toolCall(String(payload.call_id ?? payload.id), "shell", payload.action?.command);
    else if (payload.type === "function_call_output" || payload.type === "custom_tool_call_output") transcript.toolResult(String(payload.call_id), toolOutput(payload.output));
  });
  await copyRollout(homes, match.file);
  const messages = transcript.build();
  const first = messages.find((message) => message.kind === "user")?.text;
  return { providerSessionId: id, title: names.get(id) ?? titleFromText(first), startedAt: startedAt ?? match.mtime, updatedAt: updatedAt ?? match.mtime, messages };
}
