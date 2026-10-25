import { describe, it, expect } from "vitest";
import { matchesHistory, snapshotHistory } from "../gateway/src/history.js";
import { transcriptPrompt, type ChatMessage, type ContentBlock } from "../gateway/src/protocol.js";

const image = { type: "image", source: { type: "base64", media_type: "image/png", data: "OLD_IMAGE_BASE64" } };
const redaction = { type: "text", text: "[image previously shown — redacted from history]" };
function history(block: ContentBlock = image): ChatMessage[] {
  return [{ role: "user", content: [{ type: "tool_result", tool_use_id: "unique-call", content: [
    { type: "text", text: "Verified workbook state" }, block, { type: "text", text: "[id:abc123]" },
  ] }] }];
}
describe("gateway history identity", () => {
  it("accepts the exact Office image redaction without retaining image bytes", () => {
    const original = snapshotHistory(history());
    expect(matchesHistory(original, snapshotHistory(history(redaction)))).toBe(true);
    expect(JSON.stringify(original)).not.toContain("OLD_IMAGE_BASE64");
  });
  it("still rejects changed images, text, call identifiers, roles and reverse substitution", () => {
    const original = snapshotHistory(history());
    const edited = history({ ...image, source: { ...image.source, data: "CHANGED_IMAGE" } });
    expect(matchesHistory(original, snapshotHistory(edited))).toBe(false);
    const textEdit = history(redaction);
    ((textEdit[0]!.content as any[])[0].content as any[])[0].text = "Different workbook state";
    expect(matchesHistory(original, snapshotHistory(textEdit))).toBe(false);
    const idEdit = history(redaction);
    (idEdit[0]!.content as any[])[0].tool_use_id = "another-call";
    expect(matchesHistory(original, snapshotHistory(idEdit))).toBe(false);
    const roleEdit = history(redaction); roleEdit[0]!.role = "assistant";
    expect(matchesHistory(original, snapshotHistory(roleEdit))).toBe(false);
    expect(matchesHistory(snapshotHistory(history(redaction)), original)).toBe(false);
  });
  it("omits historical image bytes from text replay and states the vision limitation", () => {
    const prompt = transcriptPrompt({ model: "copilot/default", max_tokens: 50, messages: history() });
    expect(prompt).not.toContain("OLD_IMAGE_BASE64");
    expect(prompt).toContain("not available for visual verification");
    expect(prompt).toContain("Verified workbook state");
    expect(prompt).toContain("unique-call");
    expect((history()[0]!.content as any[])[0].content[1]).toEqual(image);
  });
});
