// Initializes the provider and native tool registration. It never sends a prompt.
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";
import store from "./runtime-store.cjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const [provider, command] = process.argv.slice(2);
let adapter, authClient;
try {
  if (!store.ids.includes(provider) || !command || !path.isAbsolute(command)) throw new Error("INVALID_ARGUMENTS");
  const module = await import(pathToFileURL(path.join(root, "dist", "providers", "src", provider + ".js")).href);
  const name = { claude: "ClaudeAdapter", copilot: "CopilotAdapter", codex: "CodexAdapter", opencode: "OpencodeAdapter" }[provider];
  const env = store.childEnvironment(process.env, { provider, command });
  adapter = new module[name]({ env, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  if (provider === "copilot") {
    const sdk = await import(pathToFileURL(createRequire(path.join(root, "package.json")).resolve("@github/copilot-sdk")).href);
    const { resolveCopilotPath } = await import(pathToFileURL(path.join(root, "dist", "providers", "src", "copilot-native.js")).href);
    authClient = new sdk.CopilotClient({ connection: sdk.RuntimeConnection.forStdio({ path: resolveCopilotPath(command, env), args: ["--no-custom-instructions", "--disable-builtin-mcps"] }), env, logLevel: "error" });
    await authClient.start();
    const status = await authClient.getAuthStatus();
    if (status.isAuthenticated !== true) throw new Error("authentication required");
    await authClient.stop(); authClient = undefined;
  } else {
    const detection = await adapter.detect(true);
    if (detection.status === "login_required") throw new Error("authentication required");
    if (!detection.installed || detection.status !== "ready") throw new Error("provider initialization failed");
  }
  const signal = AbortSignal.timeout(60000);
  const id = await adapter.startSession({ signal, sessionId: "preflight-" + randomUUID(), workbookConnectionId: "preflight", model: null, effort: "medium",
    requestUserInput: async () => ({ cancelled: true, answers: {} }), authorizeTool: async () => false,
    systemPrompt: "", mcp: { command: process.execPath, args: [path.join(root, "dist", "excel_mcp", "src", "cli.js")],
      bridgeUrl: "http://127.0.0.1:1", sessionId: "preflight", token: randomUUID(),
      env: { EXCEL_BRIDGE_URL: "http://127.0.0.1:1", EXCEL_SESSION_ID: "preflight", EXCEL_SESSION_TOKEN: randomUUID(), OFFICE_HOST: "excel", EXCEL_ALLOW_OFFICE_JS: "1" } } });
  await adapter.closeSession(id);
  process.stdout.write("OFFICE_PARTNER_PREFLIGHT_RESULT:" + JSON.stringify({ ok: true, phase: "native_registration" }) + "\n");
} catch (error) {
  const text = String(error);
  const code = /401|unauth|authentication required|not authenticated|not logged|sign.?in|log.?in|credentials|token.*(?:missing|expired)/i.test(text) ? "AUTH_REQUIRED"
    : /policy|organization|blocked|forbidden|403/i.test(text) ? "POLICY_DENIED"
      : /timed? ?out|timeout/i.test(text) ? "TIMEOUT" : /schema|parameter|protocol|unknown option|unrecognized/i.test(text) ? "PROTOCOL_INCOMPATIBLE" : "INITIALIZATION_FAILED";
  process.stdout.write("OFFICE_PARTNER_PREFLIGHT_RESULT:" + JSON.stringify({ ok: false, code }) + "\n"); process.exitCode = 1;
} finally { await authClient?.stop().catch(() => {}); await adapter?.dispose().catch(() => {}); }
