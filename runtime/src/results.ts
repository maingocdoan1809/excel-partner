import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import type { ToolResult } from "../../protocol/src/index.js";

/** Office results have one wire representation for MCP providers and the API gateway. */
export function toolResultToMcp(result: ToolResult): CallToolResult {
  let response: CallToolResult;
  if (!result.ok) response = { isError: true, content: [{ type: "text", text: JSON.stringify({ error: result.error }) }] };
  else {
    const data = result.data;
    const wrapped = data && typeof data === "object" && "mcpResult" in data ? (data as { mcpResult: CallToolResult }).mcpResult : undefined;
    const image = (data as { image?: { mimeType: string; data: string } } | null)?.image;
    if (wrapped) response = { ...wrapped, content: [...wrapped.content] };
    else if (image?.data && image.mimeType) {
      const { image: _image, ...rest } = data as Record<string, unknown>;
      response = { content: [{ type: "image", data: image.data, mimeType: image.mimeType }, { type: "text", text: JSON.stringify(rest) }] };
    } else response = { content: [{ type: "text", text: JSON.stringify(data ?? null) }] };
  }
  if (result.additional_user_messages?.length) response.content.push({ type: "text", text: "Additional messages sent directly by the user while this request was running. Read them before your next action:\n" + JSON.stringify({ additional_user_messages: result.additional_user_messages }) });
  return response;
}
