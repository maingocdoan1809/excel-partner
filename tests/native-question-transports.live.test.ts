import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { JsonRpcPeer } from "../providers/src/jsonrpc.js";
import { reserveOpencodePort } from "../providers/src/opencode-questions.js";
import { resolveCopilotPath } from "../providers/src/copilot-native.js";

function installed(command: string) {
  try { execFileSync(command, ["--version"], { stdio: "ignore", timeout: 5000 }); return true; }
  catch { return false; }
}

/** Real binaries, isolated state, no prompts or model inference and no paid account required. */
describe("native question transports (real CLIs, no model)", () => {
  it.skipIf(!installed(process.env.COPILOT_BIN ?? "copilot"))("connects the native Copilot SDK and registers its question handler", async () => {
    const { CopilotClient, RuntimeConnection } = await import("@github/copilot-sdk");
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), "copilot-question-handshake-"));
    const client = new CopilotClient({ connection: RuntimeConnection.forStdio({
      path: resolveCopilotPath(process.env.COPILOT_BIN ?? "copilot", process.env),
      args: ["--no-custom-instructions", "--disable-builtin-mcps"],
    }), env: { ...process.env, COPILOT_HOME: directory }, workingDirectory: directory, useLoggedInUser: false, logLevel: "error" });
    try {
      await client.start();
      const session = await client.createSession({ streaming: true, availableTools: ["ask_user"],
        onPermissionRequest: () => ({ kind: "reject" }), onUserInputRequest: async () => { throw new Error("No model was requested"); } });
      expect(session.sessionId).toBeTruthy();
      await session.disconnect();
    } finally { await client.stop(); fs.rmSync(directory, { recursive: true, force: true }); }
  }, 30_000);

  it.skipIf(!installed(process.env.OPENCODE_BIN ?? "opencode"))("connects OpenCode ACP and its authenticated native question API", async () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), "opencode-question-handshake-"));
    const port = await reserveOpencodePort();
    const password = "isolated-test-password";
    const peer = new JsonRpcPeer({ command: process.env.OPENCODE_BIN ?? "opencode",
      args: ["acp", "--hostname", "127.0.0.1", "--port", String(port)], cwd: directory, versionField: true,
      env: { ...process.env, XDG_DATA_HOME: directory, XDG_CACHE_HOME: directory, XDG_CONFIG_HOME: directory,
        OPENCODE_CONFIG_CONTENT: JSON.stringify({ plugin: [], permission: { "*": "deny", question: "allow" } }),
        OPENCODE_DISABLE_DEFAULT_PLUGINS: "1", OPENCODE_DISABLE_MODELS_FETCH: "1", OPENCODE_ENABLE_QUESTION_TOOL: "1",
        OPENCODE_SERVER_USERNAME: "excel-partner", OPENCODE_SERVER_PASSWORD: password },
    });
    try {
      const initialized = await peer.request<{ protocolVersion: number }>("initialize", { protocolVersion: 1, clientCapabilities: {}, clientInfo: { name: "excel-partner-test", version: "1" } }, 20_000);
      expect(initialized.protocolVersion).toBe(1);
      const url = `http://127.0.0.1:${port}/question?directory=${encodeURIComponent(directory)}`;
      const response = await fetch(url, { headers: { authorization: `Basic ${Buffer.from(`excel-partner:${password}`).toString("base64")}` }, signal: AbortSignal.timeout(5000) });
      expect(response.ok).toBe(true);
      expect(await response.json()).toEqual([]);
      expect((await fetch(url, { signal: AbortSignal.timeout(5000) })).status).toBe(401);
    } finally { await peer.close(); fs.rmSync(directory, { recursive: true, force: true }); }
  }, 30_000);
});
