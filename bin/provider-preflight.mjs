// Proves a provider's CLI works with this release's adapter: signed in, and a session opens with the add-in's tools registered.
// It never sends a prompt. The adapter does the work; this only builds it and reports a reason code.
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";
import store from "./runtime-store.cjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const [provider, command] = process.argv.slice(2);
let adapter;
try {
  if (!store.ids.includes(provider) || !command || !path.isAbsolute(command)) throw new Error("INVALID_ARGUMENTS");
  const { createAdapter } = await import(pathToFileURL(path.join(root, "dist", "providers", "src", "registry.js")).href);
  const env = store.childEnvironment(process.env, { provider, command });
  adapter = createAdapter(provider, { env, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  await adapter.verifyReady({
    command: process.execPath, args: [path.join(root, "dist", "excel_mcp", "src", "cli.js")],
    bridgeUrl: "http://127.0.0.1:1", sessionId: "preflight", token: randomUUID(),
    env: { EXCEL_BRIDGE_URL: "http://127.0.0.1:1", EXCEL_SESSION_ID: "preflight", EXCEL_SESSION_TOKEN: randomUUID(), OFFICE_HOST: "excel", EXCEL_ALLOW_OFFICE_JS: "1" },
  });
  process.stdout.write("OFFICE_PARTNER_PREFLIGHT_RESULT:" + JSON.stringify({ ok: true, phase: "native_registration" }) + "\n");
} catch (error) {
  const text = String(error);
  const code = /401|unauth|authentication required|not authenticated|not logged|sign.?in|log.?in|credentials|token.*(?:missing|expired)/i.test(text) ? "AUTH_REQUIRED"
    : /policy|organization|blocked|forbidden|403/i.test(text) ? "POLICY_DENIED"
      : /timed? ?out|timeout/i.test(text) ? "TIMEOUT" : /schema|parameter|protocol|unknown option|unrecognized/i.test(text) ? "PROTOCOL_INCOMPATIBLE" : "INITIALIZATION_FAILED";
  process.stdout.write("OFFICE_PARTNER_PREFLIGHT_RESULT:" + JSON.stringify({ ok: false, code }) + "\n"); process.exitCode = 1;
} finally { await adapter?.dispose().catch(() => {}); }
