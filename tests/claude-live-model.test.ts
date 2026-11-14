import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { ClaudeAdapter } from "../providers/src/claude/adapter.js";
import type { ProviderEvent } from "../protocol/src/index.js";
import type { StartSessionContext } from "../providers/src/types.js";

/** A stand-in for the CLI that speaks just enough stream-json: it answers control requests and each message with the model it was last given. */
const FAKE_CLI = `
import readline from "node:readline";
import fs from "node:fs";
const args = process.argv.slice(2);
const session = args[args.indexOf("--session-id") + 1] ?? "00000000-0000-4000-8000-000000000000";
let model = args.includes("--model") ? args[args.indexOf("--model") + 1] : "default";
const out = value => process.stdout.write(JSON.stringify(value) + "\\n");
fs.appendFileSync(process.env.FAKE_LOG, "start:" + model + ":" + args[args.indexOf("--effort") + 1] + ":" + (process.env.CLAUDE_CODE_EFFORT_LEVEL ?? "unpinned") + "\\n");
readline.createInterface({ input: process.stdin }).on("line", line => {
  const message = JSON.parse(line);
  if (message.type === "control_request") {
    if (message.request.subtype === "apply_flag_settings") fs.appendFileSync(process.env.FAKE_LOG, "effort:" + message.request.settings.effortLevel + "\\n");
    if (message.request.subtype === "set_model") { model = message.request.model ?? "default"; fs.appendFileSync(process.env.FAKE_LOG, "set_model:" + model + "\\n"); }
    out({ type: "control_response", response: { subtype: "success", request_id: message.request_id, response: {} } });
  } else if (message.type === "user") {
    out({ type: "assistant", message: { id: "m", content: [{ type: "text", text: "model=" + model }] } });
    out({ type: "result", session_id: session, is_error: false, usage: { input_tokens: 1, output_tokens: 1 } });
  }
});
`;

let folder = "";
let logFile = "";
beforeEach(() => {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "claude-live-model-"));
  logFile = path.join(folder, "log.txt");
  fs.writeFileSync(logFile, "");
  fs.writeFileSync(path.join(folder, "cli.mjs"), FAKE_CLI);
  // A command the adapter can start by itself on any platform.
  const launcher = process.platform === "win32" ? path.join(folder, "claude.cmd") : path.join(folder, "claude");
  fs.writeFileSync(launcher, process.platform === "win32" ? `@"${process.execPath}" "%~dp0cli.mjs" %*\r\n` : `#!/bin/sh\nexec "${process.execPath}" "${path.join(folder, "cli.mjs")}" "$@"\n`, { mode: 0o755 });
});
afterEach(() => fs.rmSync(folder, { recursive: true, force: true }));

function adapter() {
  const launcher = path.join(folder, process.platform === "win32" ? "claude.cmd" : "claude");
  return new ClaudeAdapter({ bridgeUrl: "", mcp: { command: "node", args: [] }, logger: { debug() {}, warn() {} },
    env: { ...process.env, CLAUDE_BIN: launcher, FAKE_LOG: logFile, EXCEL_AI_HOME: path.join(folder, "home") } });
}

function context(): StartSessionContext {
  return { workbookConnectionId: "w", sessionId: "s", signal: new AbortController().signal, model: null, effort: "medium",
    requestUserInput: async () => ({ cancelled: true, answers: {} }), authorizeTool: async () => false, systemPrompt: "",
    mcp: { command: "node", args: [], bridgeUrl: "", sessionId: "s", token: "", env: {} } };
}

async function turn(claude: ClaudeAdapter, id: string, requestId: string, model: string | null, effort: "low" | "medium" | "high" = "medium"): Promise<string> {
  let text = "";
  for await (const event of claude.sendTurn({ providerSessionId: id, requestId, prompt: "hi", selection: null, model, effort, signal: new AbortController().signal }) as AsyncIterable<ProviderEvent>) {
    if (event.type === "message_completed") text = event.text;
  }
  return text;
}

describe("Claude changes its model in the running CLI", () => {
  it("switches the model between turns without starting the CLI again", async () => {
    const claude = adapter();
    const id = await claude.startSession(context());
    expect(await turn(claude, id, "r1", null)).toBe("model=default");
    expect(await turn(claude, id, "r2", "haiku")).toBe("model=haiku");
    expect(await turn(claude, id, "r3", "haiku")).toBe("model=haiku");
    expect(await turn(claude, id, "r4", null)).toBe("model=default");
    await claude.closeSession(id);
    expect(fs.readFileSync(logFile, "utf8").trim().split("\n")).toEqual(["start:default:medium:unpinned", "set_model:haiku", "set_model:default"]);
  });

  it("changes the effort of the running CLI, which was started without a pinned effort", async () => {
    const claude = adapter();
    const id = await claude.startSession(context());
    await turn(claude, id, "r1", null, "medium");
    await turn(claude, id, "r2", null, "high");
    await turn(claude, id, "r3", null, "high");
    await turn(claude, id, "r4", null, "low");
    await claude.closeSession(id);
    expect(fs.readFileSync(logFile, "utf8").trim().split("\n")).toEqual(["start:default:medium:unpinned", "effort:high", "effort:low"]);
  });

  it("does not let the user's own CLAUDE_CODE_EFFORT_LEVEL pin the effort either", async () => {
    process.env.CLAUDE_CODE_EFFORT_LEVEL = "max";
    try {
      const claude = adapter();
      await claude.closeSession(await claude.startSession(context()));
    } finally { delete process.env.CLAUDE_CODE_EFFORT_LEVEL; }
    expect(fs.readFileSync(logFile, "utf8")).toContain("start:default:medium:unpinned");
  });

  it("declares model and effort as changeable in place, so the bridge does not restart for them", () => {
    expect(ClaudeAdapter.capabilities.liveModel).toBe(true);
    expect(ClaudeAdapter.capabilities.liveEffort).toBe(true);
  });
});
