import { expect, it } from "vitest";
import { toolResultToMcp } from "../shared/src/results.js";

it("preserves external MCP errors, images and structured results without mutating the original", () => {
  const original = { isError: true, content: [{ type: "text" as const, text: "external failure" }, { type: "image" as const, mimeType: "image/png", data: "aGVsbG8=" }], structuredContent: { reason: "failure" } };
  const converted = toolResultToMcp({ ok: true, data: { mcpResult: original }, additional_user_messages: [{ id: "followup", text: "Stop after this" }] });
  expect(converted).toMatchObject({ isError: true, structuredContent: original.structuredContent });
  expect(converted.content).toHaveLength(3);
  expect(JSON.stringify(converted.content.at(-1))).toContain("Stop after this");
  expect(original.content).toHaveLength(2);
});
it("uses the same Office image and error envelopes for MCP and gateway consumers", () => {
  expect(toolResultToMcp({ ok: true, data: { image: { mimeType: "image/png", data: "aGVsbG8=" }, sheet: "Sheet1" } })).toEqual({ content: [{ type: "image", mimeType: "image/png", data: "aGVsbG8=" }, { type: "text", text: '{"sheet":"Sheet1"}' }] });
  expect(toolResultToMcp({ ok: false, error: { code: "cancelled", message: "User declined" } })).toEqual({ isError: true, content: [{ type: "text", text: '{"error":{"code":"cancelled","message":"User declined"}}' }] });
});
