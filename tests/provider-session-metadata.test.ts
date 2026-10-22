import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { claudeTitle, codexTokenBaseline, copilotTitle } from "../providers/src/session-metadata.js";
import { parseOpencodeExport } from "../providers/src/opencode-usage.js";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function directory() { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "provider-title-")); dirs.push(dir); return dir; }

describe("provider-specific native metadata", () => {
  it("reads Copilot YAML titles, including quoted Unicode, without requiring a model call", async () => {
    const root = directory();
    const file = path.join(root, "session-state", "copilot-session", "workspace.yaml");
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, 'id: copilot-session\nname: "Đổi màu: đỏ #1"\n');
    expect(await copilotTitle("copilot-session", { COPILOT_HOME: root })).toBe("Đổi màu: đỏ #1");
    fs.writeFileSync(file, "name: >-\n  Change Shape\n  Color to Red\n");
    expect(await copilotTitle("copilot-session", { COPILOT_HOME: root })).toBe("Change Shape Color to Red");
    expect(await copilotTitle("../another-session", { COPILOT_HOME: root })).toBeUndefined();
  });

  it("reads Claude native title entries for the exact conversation, preserving the latest title", async () => {
    const root = directory(); const cwd = path.join(root, "workspace"); const id = "claude-session";
    const file = path.join(root, "projects", cwd.replace(/[^a-zA-Z0-9]/g, "-"), `${id}.jsonl`);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, [
      { type: "summary", summary: "Original Summary", sessionId: id },
      { type: "custom-title", customTitle: "Native Claude Title", sessionId: id },
      { type: "custom-title", customTitle: "Another conversation", sessionId: "other" },
    ].map(x => JSON.stringify(x)).join("\n") + '\n{"incomplete":');
    expect(await claudeTitle(id, cwd, { CLAUDE_CONFIG_DIR: root })).toBe("Native Claude Title");
    expect(await claudeTitle("unknown", cwd, { CLAUDE_CONFIG_DIR: root })).toBeUndefined();
  });

  it("loads Codex's last native odometer before resuming, ignoring the incomplete tail", async () => {
    const file = path.join(directory(), "rollout.jsonl");
    const row = (input: number) => JSON.stringify({ type: "event_msg", payload: { type: "token_count", info: {
      total_token_usage: { input_tokens: input, output_tokens: 20, cached_input_tokens: 30, total_tokens: input + 20 },
    } } });
    fs.writeFileSync(file, row(100) + "\n" + row(150) + '\n{"partial":');
    expect(await codexTokenBaseline(file)).toEqual({ inputTokens: 150, outputTokens: 20, cachedInputTokens: 30, totalTokens: 170 });
    expect(await codexTokenBaseline(undefined)).toBeUndefined();
  });
});

describe("OpenCode original accounting", () => {
  it("normalises cache/reasoning separately, instead of treating ACP's last call as a session total", () => {
    const result = parseOpencodeExport({ info: { id: "ses_original", title: "Add Total formulas and bold headers",
      tokens: { input: 28665, output: 546, reasoning: 56, cache: { read: 182412, write: 0 } } } }, "ses_original");
    expect(result).toMatchObject({ title: "Add Total formulas and bold headers", usage: {
      inputTokens: 211077, outputTokens: 602, reasoningTokens: 56, totalTokens: 211679,
      uncachedInputTokens: 28665, cachedInputTokens: 182412, cacheWriteTokens: 0, source: "opencode_export",
    } });
  });

  it("sums all assistant messages once for older exports and rejects another session", () => {
    const message = { info: { id: "msg-1", role: "assistant", tokens: {
      input: 10, output: 3, reasoning: 2, cache: { read: 20, write: 5 },
    } } };
    const data = { info: { id: "ses_original" }, messages: [message, message] };
    expect(parseOpencodeExport(data, "ses_original").usage).toMatchObject({ totalTokens: 40, outputTokens: 5 });
    expect(() => parseOpencodeExport(data, "ses_other")).toThrow("different session");
    expect(() => parseOpencodeExport({ info: { id: "ses_original" }, messages: [{ info: { id: "bad", role: "assistant" } }] }, "ses_original"))
      .toThrow("missing original usage");
  });
});
