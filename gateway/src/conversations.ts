import fs from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";
import { z } from "zod";
import type { MessagesRequest } from "./protocol.js";
import type { HistorySnapshot } from "./history.js";
import { ApiError } from "./protocol.js";

/** Claude Office serializes its historyChatId inside Anthropic metadata.user_id.
 * Auxiliary inference IDs are device/day scoped and must never identify a chat. */
export function conversationKey(request: MessagesRequest, owner: string, provider: string): string | undefined {
  const raw = request.metadata?.user_id;
  if (typeof raw !== "string" || raw.length > 8192) return;
  try {
    const id: unknown = JSON.parse(raw)?.session_id;
    if (typeof id !== "string" || !/^[a-zA-Z0-9][a-zA-Z0-9_.:-]{0,199}$/.test(id) || id.startsWith("aux:")) return;
    return createHash("sha256").update(JSON.stringify([owner, provider, id])).digest("hex");
  } catch { return; }
}

const valueBlock = z.object({ kind: z.enum(["value", "image", "redacted_image"]), hash: z.string().max(64) });
type Block = HistorySnapshot[number]["content"][number];
const block: z.ZodType<Block> = z.lazy(() => z.union([valueBlock,
  z.object({ kind: z.literal("tool_result"), hash: z.string().max(64), content: z.array(block) }),
]));
const snapshot = z.array(z.object({ role: z.string(), content: z.array(block) }));
const record = z.object({ key: z.string().regex(/^[a-f0-9]{64}$/), id: z.string(),
  nativeId: z.string().nullable(), history: snapshot.optional(),
  status: z.enum(["idle", "inflight", "blocked"]), updated: z.number() });
export type ConversationRecord = z.infer<typeof record>;

/** Synchronous atomic checkpoints keep request admission and persistence ordered.
 * No tokens, prompts, tool payloads, client IDs or image bytes are stored. */
export class ConversationStore {
  private readonly records = new Map<string, ConversationRecord>();
  constructor(private readonly file?: string) {
    if (!file || !fs.existsSync(file)) return;
    if (fs.statSync(file).size > 32 * 1024 * 1024) throw new Error("Conversation checkpoint file exceeds 32 MiB; recovery was not attempted");
    const data = z.object({ version: z.literal(1), records: z.array(record).max(512) })
      .parse(JSON.parse(fs.readFileSync(file, "utf8")));
    for (const entry of data.records) this.records.set(entry.key, entry);
  }
  get(key: string) { return this.records.get(key); }
  delete(key: string) {
    if (!this.records.has(key)) return;
    const next = new Map(this.records);
    next.delete(key);
    this.write(next);
    this.records.clear();
    for (const [k, value] of next) this.records.set(k, value);
  }
  private write(next: Map<string, ConversationRecord>, keep?: string) {
    if (!this.file) return;
    let serialized = JSON.stringify({ version: 1, records: [...next.values()] });
    while (Buffer.byteLength(serialized) > 32 * 1024 * 1024) {
      const removable = [...next.values()].find(value => value.status === "idle" && value.key !== keep);
      if (!removable) throw new ApiError(429, "rate_limit_error", "Conversation checkpoint storage limit reached");
      next.delete(removable.key);
      serialized = JSON.stringify({ version: 1, records: [...next.values()] });
    }
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    const temporary = `${this.file}.tmp`;
    fs.writeFileSync(temporary, serialized, { mode: 0o600 });
    fs.renameSync(temporary, this.file);
  }
  set(entry: ConversationRecord) {
    // A deep copy prevents a subsequent request from mutating a checkpoint.
    const next = new Map(this.records);
    next.delete(entry.key);
    next.set(entry.key, structuredClone(entry));
    if (next.size > 512) {
      const removable = [...next.values()].find(value => value.status === "idle");
      if (!removable) throw new ApiError(429, "rate_limit_error", "Conversation checkpoint limit reached");
      next.delete(removable.key);
    }
    this.write(next, entry.key);
    this.records.clear();
    for (const [key, value] of next) this.records.set(key, value);
  }
}
