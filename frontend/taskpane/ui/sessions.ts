import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { isToolNameForHost } from "../../../protocol/src/host.js";

export type { StoredSession, StoredMessage, StoredTool, StoredChange, SessionHost, SessionSummary } from "../../../protocol/src/history.js";
import type { StoredSession, StoredMessage, StoredTool, SessionHost } from "../../../protocol/src/history.js";

export function inferHost(session: Pick<StoredSession, "messages">): SessionHost {
  for (const message of session.messages) {
    for (const tool of [message.tool, ...(message.tools ?? [])]) {
      if (!tool) continue;
      if (isToolNameForHost("word", tool.name) && !isToolNameForHost("excel", tool.name)) return "word";
      if (isToolNameForHost("powerpoint", tool.name) && !isToolNameForHost("excel", tool.name)) return "powerpoint";
    }
  }
  return "excel";
}

/** Longest text kept per message or tool field, and per generated document. */
export const MAX_STORED_TEXT = 20_000;
export const MAX_STORED_ARTIFACT = 200_000;

const NEW_TITLE = "New chat";

export function newSession(provider: string, now = Date.now()): StoredSession {
  const id = `${now.toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  return { id, title: NEW_TITLE, provider, identityVersion: 1, startedAt: now, updatedAt: now, messages: [], changes: [] };
}

/** First user message, one line, as the list title. */
export function titleFrom(messages: StoredMessage[]): string {
  const first = messages.find((m) => m.kind === "user");
  if (!first) return NEW_TITLE;
  const line = first.text.replace(/\s+/g, " ").trim();
  return line.length > 70 ? `${line.slice(0, 67)}…` : line || NEW_TITLE;
}

/** Cut oversized fields so one runaway tool result cannot make a chat too large to save. Returns a copy. */
export function clipSession(session: StoredSession): StoredSession {
  const clip = (text: string, max = MAX_STORED_TEXT) => (text.length > max ? `${text.slice(0, max)}…` : text);
  const clipTool = (tool: StoredTool): StoredTool => ({ ...tool, args: clip(tool.args),
    ...(tool.input === undefined ? {} : { input: clip(tool.input) }), ...(tool.result === undefined ? {} : { result: clip(tool.result) }) });
  return {
    ...session,
    messages: session.messages.map((m) => ({ ...m, text: clip(m.text),
      ...(m.tool ? { tool: clipTool(m.tool) } : {}), ...(m.tools ? { tools: m.tools.map(clipTool) } : {}),
      ...(m.artifact ? { artifact: { ...m.artifact, content: clip(m.artifact.content, MAX_STORED_ARTIFACT) } } : {}) })),
  };
}

/** Keep each native CLI session as one history item. Rebinding the same ID resumes
 * it; a new ID moves only this request to a new item and preserves previous work.
 */
export function bindProviderSession(session: StoredSession,
  info: { provider: string; providerSessionId: string; title?: string },
  requestMessageIndex: number, requestChangeIndex: number, startedAt: number,
): { current: StoredSession; previous: StoredSession | null } {
  info = { ...info, title: cleanSessionTitle(info.title) };
  const id = `${info.provider}:${info.providerSessionId}`;
  if (session.id === id) {
    session.identityVersion = 1;
    session.providerSessionId = info.providerSessionId;
    if (info.title) { session.providerTitle = info.title; session.title = info.title; }
    return { current: session, previous: null };
  }
  const previous = { ...session, messages: session.messages.slice(0, requestMessageIndex),
    changes: session.changes.slice(0, requestChangeIndex) };
  const current: StoredSession = { ...newSession(info.provider, startedAt), id,
    providerSessionId: info.providerSessionId, providerTitle: info.title,
    messages: session.messages.slice(requestMessageIndex), changes: session.changes.slice(requestChangeIndex) };
  current.title = info.title || titleFrom(current.messages);
  return { current, previous };
}

/**
 * The earlier conversation as plain text for the model: the provider process starts empty after a reload or when a
 * past chat is continued, so the transcript (newest part, within `maxChars`) and the workbook changes are replayed once.
 */
export function buildContextPreface(session: Pick<StoredSession, "messages" | "changes">, maxChars = 8_000): string {
  const lines: string[] = [];
  let used = 0;
  const turns = session.messages.filter((m) => m.kind === "user" || m.kind === "assistant");
  for (let i = turns.length - 1; i >= 0; i--) {
    const m = turns[i]!;
    const body = m.text.length > 2_000 ? `${m.text.slice(0, 2_000)}…` : m.text;
    const who = m.kind !== "user" ? "Assistant" : m.from !== undefined ? `Request from the assistant of workbook "${m.from}"` : "User";
    const line = `${who}: ${body}`;
    if (used + line.length > maxChars && lines.length > 0) {
      lines.unshift("(earlier messages omitted)");
      break;
    }
    lines.unshift(line);
    used += line.length;
  }
  const done = session.changes.filter((c) => c.ok).map((c) => `- ${c.summary}`);
  const shown = done.slice(-20);
  return [
    '<previous_conversation note="Earlier part of this chat. You have no memory of it; use it only as context for the next request.">',
    ...lines,
    ...(shown.length > 0 ? ["", `Changes already made to the workbook${done.length > shown.length ? " (latest)" : ""}:`, ...shown] : []),
    "</previous_conversation>",
  ].join("\n");
}

/** `5 min ago`, `Yesterday 14:05`, `12 Oct 09:30`: compact timestamps for the list. */
export function formatWhen(ms: number, now = Date.now()): string {
  const diff = now - ms;
  if (diff < 60_000) return "Just now";
  if (diff < 3_600_000) return `${Math.floor(diff / 60_000)} min ago`;
  const d = new Date(ms);
  const n = new Date(now);
  const time = `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  if (d.toDateString() === n.toDateString()) return `Today ${time}`;
  const yesterday = new Date(now - 86_400_000);
  if (d.toDateString() === yesterday.toDateString()) return `Yesterday ${time}`;
  return `${d.getDate()} ${d.toLocaleString("en", { month: "short" })} ${time}`;
}
