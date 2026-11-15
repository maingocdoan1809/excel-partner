import fs from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import path from "node:path";
import { parse, stringify } from "yaml";
import type { ExternalSessionSummary, ImportedSession } from "../../../protocol/src/external-sessions.js";
import { MAX_LISTED_SESSIONS, TranscriptBuilder, isOwnWorkspace, mtimeOf, newestFirst, samePath, scanJsonLines, type SessionHomes, titleFromText, toMillis, userHome } from "../external-sessions.js";
import { nativeTitle, validNativeId } from "../session-metadata.js";
import { ensureWorkspaceDir } from "../workspace.js";

const MAX_VSCODE_FILE = 32 * 1024 * 1024;
const VSCODE_PREFIX = "vscode-chat-";

export const copilotHome = (env: NodeJS.ProcessEnv): string => env.COPILOT_HOME ?? path.join(userHome(env), ".copilot");
const stateRoot = (env: NodeJS.ProcessEnv): string => path.join(copilotHome(env), "session-state");

/** The VS Code user-data folders, stable and Insiders, where the Copilot Chat extension keeps its conversations. */
function vscodeUserDirs(env: NodeJS.ProcessEnv): string[] {
  const home = userHome(env);
  const base = process.platform === "win32" ? env.APPDATA ?? path.join(home, "AppData", "Roaming")
    : process.platform === "darwin" ? path.join(home, "Library", "Application Support") : env.XDG_CONFIG_HOME ?? path.join(home, ".config");
  return ["Code", "Code - Insiders"].map((name) => path.join(base, name, "User"));
}

// ---- Conversations of the Copilot CLI (a terminal, or VS Code's Copilot CLI sessions): ~/.copilot/session-state/<id> ----

async function readWorkspaceYaml(env: NodeJS.ProcessEnv, id: string): Promise<Record<string, any> | undefined> {
  try {
    const file = path.join(stateRoot(env), id, "workspace.yaml");
    if ((await fs.stat(file)).size > 64 * 1024) return undefined;
    const value = parse(await fs.readFile(file, "utf8"), { maxAliasCount: 0 });
    return value && typeof value === "object" ? value : undefined;
  } catch { return undefined; }
}

async function firstPrompt(file: string): Promise<string> {
  let prompt = "";
  await scanJsonLines(file, (line) => { if (line.type === "user.message" && typeof line.data?.content === "string") prompt = line.data.content; return !prompt; }, { maxBytes: 512 * 1024 });
  return prompt;
}

async function listCliSessions(env: NodeJS.ProcessEnv): Promise<ExternalSessionSummary[]> {
  let ids: string[] = [];
  try { ids = await fs.readdir(stateRoot(env)); } catch { return []; }
  const dirs: Array<{ id: string; mtime: number }> = [];
  for (const id of ids) if (validNativeId(id)) { const mtime = await mtimeOf(path.join(stateRoot(env), id, "events.jsonl")); if (mtime !== undefined) dirs.push({ id, mtime }); }
  const summaries = await Promise.all(newestFirst(dirs, MAX_LISTED_SESSIONS).map(async ({ id, mtime }): Promise<ExternalSessionSummary | null> => {
    const meta = await readWorkspaceYaml(env, id);
    const cwd = typeof meta?.cwd === "string" ? meta.cwd : undefined;
    if (isOwnWorkspace(cwd)) return null;
    const prompt = await firstPrompt(path.join(stateRoot(env), id, "events.jsonl")).catch(() => "");
    if (!prompt) return null;
    const client = String(meta?.client_name ?? "");
    return { id, title: nativeTitle(meta?.name) ?? titleFromText(prompt), origin: /vscode|visual studio code/i.test(client) ? "VS Code" : "Terminal",
      ...(cwd ? { cwd } : {}), updatedAt: toMillis(meta?.updated_at) ?? mtime };
  }).map((pending) => pending.catch(() => null)));
  return summaries.filter((item): item is ExternalSessionSummary => item !== null);
}

/** The CLI loads a conversation from its own state folder: a private one gets a copy of the conversation's folder. */
async function copySessionFolder(homes: SessionHomes, id: string): Promise<void> {
  const from = stateRoot(homes.source), to = stateRoot(homes.target);
  if (samePath(from, to)) return;
  const destination = path.join(to, id);
  try { await fs.access(destination); return; } catch { /* Not copied yet. */ }
  await fs.mkdir(to, { recursive: true });
  const staging = `${destination}.importing-${randomUUID()}`;
  try { await fs.cp(path.join(from, id), staging, { recursive: true }); await fs.rename(staging, destination); }
  catch (error) { await fs.rm(staging, { recursive: true, force: true }); throw error; }
}

async function importCliSession(homes: SessionHomes, id: string): Promise<ImportedSession> {
  const env = homes.source;
  const file = path.join(stateRoot(env), id, "events.jsonl");
  const mtime = await mtimeOf(file);
  if (mtime === undefined) throw new Error("That Copilot conversation was not found on this computer.");
  const meta = await readWorkspaceYaml(env, id);
  const transcript = new TranscriptBuilder();
  let startedAt: number | undefined, updatedAt: number | undefined;
  await scanJsonLines(file, (line) => {
    const at = toMillis(line.timestamp);
    if (at !== undefined) { startedAt ??= at; updatedAt = at; }
    const data = line.data;
    if (!data) return;
    if (line.type === "user.message") transcript.user(data.content);
    else if (line.type === "assistant.message") {
      transcript.assistant(data.content);
      for (const request of Array.isArray(data.toolRequests) ? data.toolRequests : []) transcript.toolCall(String(request.toolCallId), String(request.toolTitle ?? request.name ?? "tool"), request.arguments);
    } else if (line.type === "tool.execution_complete") {
      const output = typeof data.result?.content === "string" ? data.result.content : typeof data.error?.message === "string" ? data.error.message : undefined;
      transcript.toolResult(String(data.toolCallId), output, data.success === false);
    }
  });
  await copySessionFolder(homes, id);
  const messages = transcript.build();
  const first = messages.find((message) => message.kind === "user")?.text;
  return { providerSessionId: id, title: nativeTitle(meta?.name) ?? titleFromText(first), startedAt: toMillis(meta?.created_at) ?? startedAt ?? mtime, updatedAt: updatedAt ?? mtime, messages };
}

// ---- Conversations of the Copilot Chat extension in VS Code: <user>/workspaceStorage/*/chatSessions/<id>.jsonl ----

interface ChatFile { file: string; sessionId: string; mtime: number; folder?: string }

async function chatFiles(env: NodeJS.ProcessEnv): Promise<ChatFile[]> {
  const found: ChatFile[] = [];
  const collect = async (dir: string, folder?: string) => {
    let names: string[] = [];
    try { names = await fs.readdir(dir); } catch { return; }
    for (const name of names) {
      const match = /^(.+)\.jsonl?$/.exec(name);
      if (!match || !validNativeId(match[1]!)) continue;
      const file = path.join(dir, name);
      const mtime = await mtimeOf(file);
      if (mtime !== undefined) found.push({ file, sessionId: match[1]!, mtime, folder });
    }
  };
  for (const user of vscodeUserDirs(env)) {
    await collect(path.join(user, "globalStorage", "emptyWindowChatSessions"));
    let workspaces: string[] = [];
    try { workspaces = await fs.readdir(path.join(user, "workspaceStorage")); } catch { continue; }
    for (const hash of workspaces) {
      const root = path.join(user, "workspaceStorage", hash);
      await collect(path.join(root, "chatSessions"), await workspaceFolder(root));
    }
  }
  return found;
}

/** The folder a VS Code workspace was opened on, from its `workspace.json` (`file:///c%3A/Projects/app`). */
async function workspaceFolder(root: string): Promise<string | undefined> {
  try {
    const value = JSON.parse(await fs.readFile(path.join(root, "workspace.json"), "utf8"));
    const uri = typeof value.folder === "string" ? value.folder : typeof value.workspace === "string" ? value.workspace : undefined;
    return uri?.startsWith("file:") ? decodeURIComponent(new URL(uri).pathname).replace(/^\/(?=[a-zA-Z]:)/, "") : undefined;
  } catch { return undefined; }
}

/** Applies one VS Code mutation-log entry: 0 = whole state, 1 = set, 2 = push (with `i`: cut the array there first), 3 = delete. */
function applyEntry(state: any, entry: any): any {
  if (entry.kind === 0) return entry.v;
  if (!state || !Array.isArray(entry.k)) return state;
  const keys: Array<string | number> = entry.k;
  let target = state;
  for (const key of keys.slice(0, entry.kind === 2 ? keys.length : -1)) { target = target?.[key]; if (target === undefined || target === null) return state; }
  if (entry.kind === 2) {
    if (!Array.isArray(target)) return state;
    if (typeof entry.i === "number") target.length = Math.min(target.length, entry.i);
    if (Array.isArray(entry.v)) target.push(...entry.v);
  } else if (entry.kind === 1) target[keys.at(-1)!] = entry.v;
  else if (entry.kind === 3) delete target[keys.at(-1)!];
  return state;
}

/** One saved conversation: a log of changes (`.jsonl`), or the whole object (`.json`, older versions). */
async function readChat(file: string): Promise<any | undefined> {
  const stat = await fs.stat(file);
  if (stat.size > MAX_VSCODE_FILE) return undefined;
  const text = await fs.readFile(file, "utf8");
  if (file.endsWith(".json")) { try { return JSON.parse(text); } catch { return undefined; } }
  let state: any;
  for (const line of text.split("\n")) {
    if (!line.trim()) continue;
    try { state = applyEntry(state, JSON.parse(line)); } catch { /* A damaged line loses only its own change. */ }
  }
  return state;
}

const requestText = (request: any): string => typeof request?.message === "string" ? request.message : typeof request?.message?.text === "string" ? request.message.text : "";

async function listChatSessions(env: NodeJS.ProcessEnv): Promise<ExternalSessionSummary[]> {
  const summaries = await Promise.all(newestFirst(await chatFiles(env), MAX_LISTED_SESSIONS).map(async (item): Promise<ExternalSessionSummary | null> => {
    const chat = await readChat(item.file);
    const requests: any[] = Array.isArray(chat?.requests) ? chat.requests.filter((request: any) => requestText(request).trim()) : [];
    // A chat panel that was opened and never used is saved too.
    if (!requests.length) return null;
    const first = requests[0];
    return { id: VSCODE_PREFIX + item.sessionId, title: nativeTitle(chat.customTitle) ?? titleFromText(requestText(first)), origin: "VS Code",
      ...(item.folder ? { cwd: item.folder } : {}), updatedAt: toMillis(requests.at(-1).timestamp) ?? item.mtime, messageCount: requests.length };
  }).map((pending) => pending.catch(() => null)));
  return summaries.filter((item): item is ExternalSessionSummary => item !== null);
}

/** The text of a response: its markdown, with file references written as `name`; thinking and edit blocks are left out. */
function responseText(parts: unknown): string {
  if (!Array.isArray(parts)) return "";
  return parts.map((part: any) => {
    if (typeof part?.value === "string" && !part.kind) return part.value;
    if (part?.kind === "markdownContent" && typeof part.content?.value === "string") return part.content.value;
    if (part?.kind === "inlineReference") {
      const name = part.name ?? (typeof part.inlineReference?.path === "string" ? path.posix.basename(part.inlineReference.path) : undefined);
      return name ? `\`${name}\`` : "";
    }
    return "";
  }).join("");
}

const toolText = (value: unknown): string => typeof value === "string" ? value : typeof (value as any)?.value === "string" ? (value as any).value : "";

/** A fixed id for a VS Code conversation, so importing it twice opens the same conversation. */
function nativeIdFor(sessionId: string): string {
  const hex = createHash("sha256").update(`vscode-chat:${sessionId}`).digest("hex");
  // The CLI only accepts a well-formed version 4 id.
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-${"89ab"[parseInt(hex[16]!, 16) & 3]}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

/**
 * The Copilot Chat extension's conversations are not Copilot CLI sessions, so the CLI cannot resume them. The CLI keeps
 * a conversation as a folder of events that it replays on `session/load`, so one is written in that form from the
 * chat's questions and answers. It is written only once: later turns made here are the CLI's own.
 */
async function writeNativeSession(env: NodeJS.ProcessEnv, cliVersion: string, nativeId: string, title: string, turns: Array<{ user: string; assistant: string; at: number }>): Promise<void> {
  const dir = path.join(stateRoot(env), nativeId);
  try { await fs.access(dir); return; } catch { /* Not written yet. */ }
  const cwd = ensureWorkspaceDir();
  const events: unknown[] = [];
  let parent: string | null = null;
  const add = (type: string, data: unknown, at: number) => {
    const id = randomUUID();
    events.push({ type, data, id, timestamp: new Date(at).toISOString(), parentId: parent });
    parent = id;
  };
  const startedAt = turns[0]?.at ?? Date.now();
  add("session.start", { sessionId: nativeId, version: 1, producer: "copilot-agent", copilotVersion: cliVersion, startTime: new Date(startedAt).toISOString(), context: { cwd }, alreadyInUse: false, remoteSteerable: false }, startedAt);
  for (const turn of turns) {
    const messageId = randomUUID(), interactionId = randomUUID();
    add("user.message", { content: turn.user, messageId, interactionId }, turn.at);
    add("assistant.turn_start", { turnId: "0", interactionId }, turn.at);
    add("assistant.message", { messageId: randomUUID(), originatingMessageId: messageId, content: turn.assistant, toolRequests: [], interactionId, turnId: "0" }, turn.at);
    add("assistant.turn_end", { turnId: "0" }, turn.at);
  }
  const endedAt = turns.at(-1)?.at ?? startedAt;
  // Written beside the final folder and renamed into place, so the CLI never sees a half-written conversation.
  const staging = `${dir}.importing-${randomUUID()}`;
  await fs.mkdir(staging, { recursive: true });
  try {
    await fs.writeFile(path.join(staging, "events.jsonl"), events.map((event) => JSON.stringify(event)).join("\n") + "\n");
    await fs.writeFile(path.join(staging, "workspace.yaml"), stringify({
      id: nativeId, cwd, client_name: "github/cli", last_client_name: "github/cli", name: title, user_named: false, summary_count: 0, fork_count: 0,
      created_at: new Date(startedAt).toISOString(), updated_at: new Date(endedAt).toISOString(),
    }));
    await fs.rename(staging, dir);
  } catch (error) { await fs.rm(staging, { recursive: true, force: true }); throw error; }
}

async function importChatSession(homes: SessionHomes, cliVersion: string, sessionId: string): Promise<ImportedSession> {
  const item = (await chatFiles(homes.source)).find((candidate) => candidate.sessionId === sessionId);
  const chat = item ? await readChat(item.file) : undefined;
  if (!item || !chat) throw new Error("That VS Code conversation was not found on this computer.");
  const transcript = new TranscriptBuilder();
  const turns: Array<{ user: string; assistant: string; at: number }> = [];
  for (const request of Array.isArray(chat.requests) ? chat.requests : []) {
    const user = requestText(request).trim();
    if (!user) continue;
    const at = toMillis(request.timestamp) ?? item.mtime;
    transcript.user(user);
    let reply = "";
    for (const part of Array.isArray(request.response) ? request.response : []) {
      if (part?.kind === "toolInvocationSerialized") {
        transcript.assistant(reply); reply = "";
        const name = String(part.toolId ?? "tool");
        transcript.toolCall(String(part.toolCallId ?? name), name, toolText(part.invocationMessage));
        transcript.toolResult(String(part.toolCallId ?? name), toolText(part.pastTenseMessage), Boolean(part.resultError));
      } else reply += responseText([part]);
    }
    transcript.assistant(reply);
    const answer = responseText(request.response).trim();
    if (answer) turns.push({ user, assistant: answer, at });
  }
  if (!turns.length) throw new Error("That VS Code conversation has no answers to continue from.");
  const messages = transcript.build();
  const title = nativeTitle(chat.customTitle) ?? titleFromText(messages.find((message) => message.kind === "user")?.text);
  const nativeId = nativeIdFor(sessionId);
  await writeNativeSession(homes.target, cliVersion, nativeId, title, turns);
  return { providerSessionId: nativeId, title, startedAt: toMillis(chat.creationDate) ?? item.mtime, updatedAt: turns.at(-1)!.at, messages };
}

// ---- What the adapter offers ----

export async function listCopilotSessions(env: NodeJS.ProcessEnv): Promise<ExternalSessionSummary[]> {
  const lists = await Promise.all([listCliSessions(env), listChatSessions(env)]);
  return lists.flat().sort((a, b) => b.updatedAt - a.updatedAt).slice(0, MAX_LISTED_SESSIONS);
}

/** `cliVersion` is the installed Copilot CLI's: it refuses a conversation whose first event does not name the version that wrote it. */
export async function importCopilotSession(homes: SessionHomes, id: string, cliVersion = "1.0.0"): Promise<ImportedSession> {
  if (id.startsWith(VSCODE_PREFIX)) {
    const sessionId = id.slice(VSCODE_PREFIX.length);
    if (!validNativeId(sessionId)) throw new Error("Invalid VS Code conversation ID");
    return importChatSession(homes, cliVersion, sessionId);
  }
  if (!validNativeId(id)) throw new Error("Invalid Copilot session ID");
  return importCliSession(homes, id);
}
