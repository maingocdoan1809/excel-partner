import type { ContentBlock } from "./protocol.js";
import { CallToolResultSchema, type CallToolResult } from "@modelcontextprotocol/sdk/types.js";

export const CONVERSION_ERROR = "agent-gateway/conversion-error";

function conversionFailure(blocks: ContentBlock[], reason: string): CallToolResult {
  return {
    content: [
      ...blocks.filter((block): block is { type: "text"; text: string } => block?.type === "text" && typeof block.text === "string")
        .map(block => ({ type: "text" as const, text: block.text })),
      { type: "text", text: `Gateway could not forward the client's tool result: ${reason}. The client operation may already have executed. Inspect its state before repeating a write.` },
    ],
    isError: true,
    _meta: { [CONVERSION_ERROR]: true },
  };
}

/** Preserve the client's content/error marker and also supply MCP's structured
 * result when possible. CLI code tools must not have to guess JSON vs text. */
export function clientToolResult(content: unknown, isError: boolean): CallToolResult {
  const blocks: ContentBlock[] = typeof content === "string" ? [{ type: "text", text: content }]
    : Array.isArray(content) ? content : [{ type: "text", text: "" }];
  const converted: ContentBlock[] = [];
  for (const block of blocks) {
    if (block?.type === "image" && block.source !== undefined) {
      const source = block.source as { type?: unknown; data?: unknown; media_type?: unknown } | null;
      if (!source || source.type !== "base64" || typeof source.data !== "string" || !source.data
        || typeof source.media_type !== "string" || !source.media_type.startsWith("image/")) {
        return conversionFailure(blocks, "images must contain inline base64 data and an image media_type");
      }
      // Anthropic tool results use source.{data,media_type}; MCP uses data/mimeType.
      // Never change the original API history: session matching uses its wire form.
      converted.push({ type: "image", data: source.data, mimeType: source.media_type });
    } else converted.push(block);
  }
  const parsed = CallToolResultSchema.safeParse({ content: converted, isError });
  if (!parsed.success) return conversionFailure(blocks, "unsupported or malformed MCP content");
  const result = parsed.data;
  let structured: Record<string, unknown> | undefined;
  const texts = result.content.filter(block => block.type === "text");
  if (texts.length === 1 && result.content.every(block => block.type === "text" || block.type === "image")) {
    try {
      const value: unknown = JSON.parse(texts[0]!.text);
      if (value && typeof value === "object" && !Array.isArray(value)) structured = value as Record<string, unknown>;
    } catch { /* Non-JSON text remains a valid tool result. */ }
  }
  return { ...result, ...(structured ? { structuredContent: structured } : {}) };
}
