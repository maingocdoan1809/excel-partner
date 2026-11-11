import { z } from "zod";

/** The Office application a chat belongs to. */
export type SessionHost = "excel" | "word" | "powerpoint";

/**
 * Shared durable history contract. SQLite is owned by the local bridge; UI
 * preferences and pending recovery copies stay in browser storage.
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
  artifact?: { name: string; content: string };
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

export interface HistorySummary extends SessionSummary { host: SessionHost; revision: number }
export interface HistoryPage { items: HistorySummary[]; total: number; generation: number }
export interface HistoryRecord { session: StoredSession; revision: number; generation: number }
export interface HistoryTranscript extends HistoryRecord { totalMessages: number; totalChanges: number }
export interface HistorySave { session: StoredSession; baseRevision: number; operationId: string; generation: number }
export interface HistoryAck { revision: number; generation: number }
export interface HistoryStats { count: number; bytes: number; generation: number }
const tool = z.object({ name: z.string(), args: z.string(), input: z.string().optional(), result: z.string().optional(), state: z.enum(["running", "ok", "error"]), error: z.string().nullable() }).passthrough();
export const StoredSessionSchema = z.object({
  id: z.string().min(1).max(256), title: z.string().max(1000), provider: z.enum(["claude", "copilot", "codex", "opencode"]),
  host: z.enum(["excel", "word", "powerpoint"]).optional(), providerSessionId: z.string().max(512).optional(), providerTitle: z.string().optional(), identityVersion: z.literal(1).optional(),
  startedAt: z.number().finite(), updatedAt: z.number().finite(),
  messages: z.array(z.object({ kind: z.enum(["user", "assistant", "error", "note", "tool", "tool_group", "artifact"]), text: z.string(), tool: tool.optional(), tools: z.array(tool).optional(), artifact: z.object({ name: z.string(), content: z.string() }).optional() }).passthrough()),
  changes: z.array(z.object({ at: z.number().finite(), tool: z.string(), ok: z.boolean(), summary: z.string(), target: z.object({ sheet: z.string(), address: z.string() }).nullable() }).passthrough()),
}).passthrough();
export const HistorySaveSchema = z.object({ session: StoredSessionSchema, baseRevision: z.number().int().nonnegative(), generation: z.number().int().nonnegative(), operationId: z.string().min(1).max(128) });
