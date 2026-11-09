import type { StoredArtifact } from "./artifact.js";
import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { isToolNameForHost } from "../../../protocol/src/host.js";

/** The Office application a chat belongs to. */
export type SessionHost = "excel" | "word" | "powerpoint";

/**
 * Chat history kept in the task pane's own storage (localStorage on the add-in origin): nothing leaves the machine
 * and nothing is stored on the bridge. A "session" is one conversation, from "New chat" to the next "New chat".
 */

export interface StoredTool {
  name: string;
  args: string;
  input?: string;
  result?: string;
  state: "running" | "ok" | "error";
  error: string | null;
}

export interface StoredMessage {
  kind: "user" | "assistant" | "error" | "note" | "tool" | "tool_group" | "artifact";
  text: string;
  delivery?: "pending" | "held";
  queueId?: string;
  attachments?: Array<{ name: string }>;
  /** Set on a request another open workbook's assistant sent: the name of that workbook. `text` is the request itself. */
  from?: string;
  /** The application the request came from; messages saved before Word and PowerPoint have none: Excel. */
  fromHost?: "excel" | "word" | "powerpoint";
  /** The assistant used the connected-file tools before writing this reply, so its bare cell references could be another file's. */
  aboutOtherFiles?: boolean;
  tool?: StoredTool;
  tools?: StoredTool[];
  artifact?: StoredArtifact;
}

export interface StoredChange {
  at: number;
  tool: string;
  ok: boolean;
  summary: string;
  target: { sheet: string; address: string } | null;
  /** Ties the entry to the undo journal; the journal is memory-only, so older entries cannot be undone. */
  callId?: string;
  requestId?: string;
}

export interface StoredSession {
  id: string;
  title: string;
  providerSessionId?: string;
  providerTitle?: string;
  /** Distinguishes old chats with no recorded native identity from a new provisional chat. */
  identityVersion?: 1;
  provider: string;
  /** The application the chat was held in. History shows a chat only in the same app; chats saved before this existed have none (see `inferHost`). */
  host?: SessionHost;
  startedAt: number;
  updatedAt: number;
  messages: StoredMessage[];
  changes: StoredChange[];
}

/**
 * Which app an older chat (saved before chats recorded their app) belongs to: a Word or PowerPoint chat shows itself by
 * using a tool only that app has. A chat that never used one is taken to be Excel's, as every chat used to be.
 */
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

export interface SessionSummary {
  id: string;
  title: string;
  provider: string;
  providerSessionId?: string;
  startedAt: number;
  updatedAt: number;
  messageCount: number;
  changeCount: number;
}

type KeyValueStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;

export const LIMITS = {
  sessions: 40,
  messagesPerSession: 400,
  textChars: 20_000,
  changesPerSession: 300,
  /** Total serialised size kept (characters); oldest sessions go first. */
  totalChars: 3_000_000,
} as const;

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

/** Cap sizes so one runaway chat cannot fill the storage quota. */
export function trimSession(session: StoredSession): StoredSession {
  const clip = (text: string) => (text.length > LIMITS.textChars ? `${text.slice(0, LIMITS.textChars)}…` : text);
  return {
    ...session,
    providerTitle: cleanSessionTitle(session.provider, session.providerTitle),
    title: cleanSessionTitle(session.provider, session.providerTitle) || titleFrom(session.messages),
    messages: session.messages.slice(-LIMITS.messagesPerSession).map((m) => ({ ...m, text: clip(m.text) })),
    changes: session.changes.slice(-LIMITS.changesPerSession),
  };
}

/** Keep each native CLI session as one history item. Rebinding the same ID resumes
 * it; a new ID moves only this request to a new item and preserves previous work.
 */
export function bindProviderSession(session: StoredSession,
  info: { provider: string; providerSessionId: string; title?: string },
  requestMessageIndex: number, requestChangeIndex: number, startedAt: number,
): { current: StoredSession; previous: StoredSession | null } {
  info = { ...info, title: cleanSessionTitle(info.provider, info.title) };
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
 * All the chats on this computer live in one place (the add-in's address is the same in Word, Excel and PowerPoint), but each
 * app sees only its own: `scope` says which app this pane is in. Saving, deleting and clearing touch that app's chats only.
 */
export class SessionStore {
  constructor(
    private readonly storage: KeyValueStorage,
    private readonly key = "excelPartner.sessions.v1",
    private readonly scope: () => SessionHost = () => "excel",
  ) {}

  private read(): StoredSession[] {
    try {
      const parsed = JSON.parse(this.storage.getItem(this.key) ?? "[]") as unknown;
      return Array.isArray(parsed)
        ? (parsed as StoredSession[])
          .filter((s) => s && typeof s.id === "string" && Array.isArray(s.messages))
          .map((s) => {
            const providerTitle = cleanSessionTitle(s.provider, s.providerTitle);
            return { ...s, host: s.host ?? inferHost(s), providerTitle, title: providerTitle || (s.providerTitle ? titleFrom(s.messages) : s.title),
              messages: s.messages.filter((m) => !(m.kind === "note" && m.text === "Updating Excel Partner…")) };
          })
        : [];
    } catch {
      return []; // corrupt storage must never break the pane
    }
  }

  private write(sessions: StoredSession[]): void {
    // The count limit is per app, so a long Excel history cannot push out the chats of Word or PowerPoint.
    const perHost = new Map<string, number>();
    let kept = sessions.sort((a, b) => b.updatedAt - a.updatedAt).filter((s) => {
      const count = (perHost.get(s.host ?? "excel") ?? 0) + 1;
      perHost.set(s.host ?? "excel", count);
      return count <= LIMITS.sessions;
    });
    let json = JSON.stringify(kept);
    while (json.length > LIMITS.totalChars && kept.length > 1) {
      kept = kept.slice(0, -1);
      json = JSON.stringify(kept);
    }
    try {
      this.storage.setItem(this.key, json);
    } catch {
      // Quota exceeded: keep only the newest session rather than losing the write silently.
      try {
        this.storage.setItem(this.key, JSON.stringify(kept.slice(0, 1)));
      } catch {
        /* storage unavailable: history is best-effort */
      }
    }
  }

  /** The chats of this pane's app. */
  private mine(): StoredSession[] {
    const host = this.scope();
    return this.read().filter((s) => s.host === host);
  }

  /** Newest first; empty chats are not listed. */
  list(): SessionSummary[] {
    return this.mine()
      .filter((s) => s.messages.some((m) => m.kind === "user"))
      .sort((a, b) => b.updatedAt - a.updatedAt)
      .map((s) => ({
        id: s.id,
        title: s.title,
        provider: s.provider,
        providerSessionId: s.providerSessionId,
        startedAt: s.startedAt,
        updatedAt: s.updatedAt,
        messageCount: s.messages.filter((m) => m.kind === "user" || m.kind === "assistant").length,
        changeCount: s.changes.filter((c) => c.ok).length,
      }));
  }

  get(id: string): StoredSession | null {
    return this.mine().find((s) => s.id === id) ?? null;
  }

  /** The most recently updated non-empty session, if it is newer than `maxAgeMs`. */
  latest(maxAgeMs: number, now = Date.now()): StoredSession | null {
    const [top] = this.mine()
      .filter((s) => s.messages.some((m) => m.kind === "user"))
      .sort((a, b) => b.updatedAt - a.updatedAt);
    return top && now - top.updatedAt <= maxAgeMs ? top : null;
  }

  save(session: StoredSession): void {
    const trimmed: StoredSession = { ...trimSession(session), host: this.scope() };
    const others = this.read().filter((s) => s.id !== trimmed.id);
    // Empty chats are not worth keeping.
    this.write(trimmed.messages.some((m) => m.kind === "user") ? [trimmed, ...others] : others);
  }

  delete(id: string): void {
    const host = this.scope();
    this.write(this.read().filter((s) => !(s.id === id && s.host === host)));
  }

  /** Delete this app's chats. The chats of the other apps stay. */
  clear(): void {
    const host = this.scope();
    const rest = this.read().filter((s) => s.host !== host);
    if (rest.length === 0) this.storage.removeItem(this.key);
    else this.write(rest);
  }
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
