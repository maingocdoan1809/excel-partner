import { describe, it, expect } from "vitest";
import { conversationKey, ConversationStore } from "../gateway/src/conversations.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { MessagesRequest } from "../gateway/src/protocol.js";

const request = (user_id: unknown): MessagesRequest => ({ model: "copilot/default", max_tokens: 20,
  messages: [{ role: "user", content: "hello" }], metadata: { user_id } });
describe("Office conversation identity", () => {
  it("extracts only the real session ID and namespaces it by token and provider", () => {
    const first = request(JSON.stringify({ session_id: "history-chat-uuid", user: "account", device_id: "device-a" }));
    const key = conversationKey(first, "token-digest", "copilot");
    expect(key).toMatch(/^[a-f0-9]{64}$/);
    expect(conversationKey(request(JSON.stringify({ session_id: "history-chat-uuid", device_id: "device-b" })), "token-digest", "copilot")).toBe(key);
    expect(conversationKey(first, "other-token", "copilot")).not.toBe(key);
    expect(conversationKey(first, "token-digest", "codex")).not.toBe(key);
  });
  it.each([undefined, null, {}, "plain-user-id", "{bad json", "null", "[]", JSON.stringify({ device_id: "device-only" }),
    JSON.stringify({ session_id: "aux:visual-check:device:2026-10-24" }), JSON.stringify({ session_id: "" }),
    JSON.stringify({ session_id: "x".repeat(201) }), JSON.stringify({ session_id: 12 })])("does not guess chat identity from invalid or auxiliary input (%j)", raw => {
    expect(conversationKey(request(raw), "owner", "copilot")).toBeUndefined();
  });
  it("fails closed on invalid checkpoints instead of silently losing native identity", () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), "gateway-invalid-checkpoint-"));
    try {
      const file = path.join(directory, "conversations.json");
      for (const invalid of ["{broken", JSON.stringify({ version: 2, records: [] }), JSON.stringify({ version: 1, records: [{ key: "invalid" }] })]) {
        fs.writeFileSync(file, invalid);
        expect(() => new ConversationStore(file)).toThrow();
        expect(fs.readFileSync(file, "utf8")).toBe(invalid);
      }
    } finally { fs.rmSync(directory, { recursive: true, force: true }); }
  });
});
