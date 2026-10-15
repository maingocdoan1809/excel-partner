/**
 * Call an Excel tool on a LIVE task pane through the bridge, exactly like an MCP
 * client would. Needs the bridge started with BRIDGE_DEV=1.
 *
 *   BRIDGE_DEV=1 bash scripts/start.sh
 *   npx tsx scripts/tool-call.ts get_selection
 *   npx tsx scripts/tool-call.ts add_shape '{"type":"FlowChartProcess","anchor":"M10:M11","name":"a"}'
 */
import https from "node:https";

const [tool, rawArgs = "{}"] = process.argv.slice(2);
if (!tool) {
  console.error("usage: tool-call.ts <tool> [json-args]");
  process.exit(2);
}
const base = process.env.BRIDGE_URL ?? "https://127.0.0.1:8765";
const agent = new https.Agent({ rejectUnauthorized: false }); // loopback dev bridge, self-signed

function request(method: string, path: string, body?: unknown, token?: string): Promise<any> {
  return new Promise((resolve, reject) => {
    const req = https.request(
      new URL(path, base),
      {
        method,
        agent,
        headers: {
          "content-type": "application/json",
          ...(token ? { authorization: `Bearer ${token}` } : {}),
        },
      },
      (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => {
          try {
            resolve(JSON.parse(data));
          } catch {
            reject(new Error(`HTTP ${res.statusCode}: ${data.slice(0, 200)}`));
          }
        });
      },
    );
    req.on("error", reject);
    req.end(body === undefined ? undefined : JSON.stringify(body));
  });
}

const sessions: Array<{ id: string; token: string; connected: boolean }> = await request("GET", "/debug/sessions");
const live = sessions.filter((s) => s.connected);
if (live.length === 0) {
  console.error("no connected task pane (open Excel > Home > Excel Partner), or the bridge lacks BRIDGE_DEV=1");
  process.exit(1);
}
const session = live.at(-1)!;
const result = await request("POST", "/internal/tool", { sessionId: session.id, tool, args: JSON.parse(rawArgs) }, session.token);
console.log(JSON.stringify(result, null, 2));
process.exit(result.ok ? 0 : 1);
