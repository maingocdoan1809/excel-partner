import { createHash } from "node:crypto";
import { stable, type ChatMessage, type ContentBlock } from "./protocol.js";

const REDACTED_IMAGE = "[image previously shown — redacted from history]";
const hash = (value: unknown) => createHash("sha256").update(stable(value)).digest("hex");
type Block = { kind: "value" | "image" | "redacted_image"; hash: string }
  | { kind: "tool_result"; hash: string; content: Block[] };
export type HistorySnapshot = Array<{ role: string; content: Block[] }>;

/** Keep only digests, not old image bytes. Office redacts images from older
 * tool results; that one directional replacement does not change the chat. */
export function snapshotHistory(messages: ChatMessage[]): HistorySnapshot {
  function blocks(content: string | ContentBlock[]): Block[] {
    const result: Block[] = [];
    let text = "";
    let hasText = false;
    const flush = () => { if (hasText) result.push({ kind: "value", hash: hash({ type: "text", text }) }); text = ""; hasText = false; };
    for (const block of typeof content === "string" ? [{ type: "text", text: content }] : content) {
      const plain = block.type === "text" && typeof block.text === "string"
        && Object.keys(block).every(key => ["type", "text", "cache_control"].includes(key));
      if (plain && block.text !== REDACTED_IMAGE) { text += block.text; hasText = true; continue; }
      flush();
      if (plain && block.text === REDACTED_IMAGE) result.push({ kind: "redacted_image", hash: "" });
      else if (block.type === "image") result.push({ kind: "image", hash: hash(block) });
      else if (block.type === "tool_result" && (typeof block.content === "string" || Array.isArray(block.content))) {
        const { content: nested, ...fields } = block;
        result.push({ kind: "tool_result", hash: hash(fields), content: blocks(nested as string | ContentBlock[]) });
      } else result.push({ kind: "value", hash: hash(block) });
    }
    flush();
    return result;
  }
  return messages.map(message => ({ role: message.role, content: blocks(message.content) }));
}

export function matchesHistory(expected: HistorySnapshot, actual: HistorySnapshot): boolean {
  function match(a: Block[], b: Block[]): boolean {
    return a.length === b.length && a.every((left, index) => {
      const right = b[index]!;
      if (left.kind === "image" && right.kind === "redacted_image") return true;
      if (left.kind !== right.kind || left.hash !== right.hash) return false;
      return left.kind !== "tool_result" || (right.kind === "tool_result" && match(left.content, right.content));
    });
  }
  return expected.length === actual.length && expected.every((message, index) =>
    message.role === actual[index]!.role && match(message.content, actual[index]!.content));
}
