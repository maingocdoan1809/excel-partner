import { gatewayToolInstructions } from "./tool-instructions.js";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { ListToolsRequestSchema, CallToolRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import http from "node:http";
import https from "node:https";

const base = process.env.AGENT_GATEWAY_RELAY_URL;
const token = process.env.AGENT_GATEWAY_RELAY_TOKEN;
const session = process.env.AGENT_GATEWAY_RELAY_SESSION;
if (!base || !token || !session) throw new Error("Missing gateway relay configuration");
async function call(route: string, body: unknown) {
  const url = new URL(`${base}${route}`);
  return new Promise<any>((resolve, reject) => {
    const req = (url.protocol === "https:" ? https : http).request(url, { method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json" } }, res => {
      const chunks: Buffer[] = []; let bytes = 0;
      res.on("data", chunk => { bytes += chunk.length; if (bytes > 16 * 1024 * 1024) res.destroy(new Error("Tool result exceeds 16 MB")); else chunks.push(Buffer.from(chunk)); });
      res.on("error", reject);
      res.on("end", () => {
        const text = Buffer.concat(chunks).toString("utf8");
        if (res.statusCode !== 200) { reject(new Error(`Gateway relay returned ${res.statusCode}: ${text.slice(0, 1000)}`)); return; }
        try { resolve(JSON.parse(text)); } catch { reject(new Error("Invalid relay JSON response")); }
      });
    });
    req.on("error", reject);
    req.setTimeout(65 * 60_000, () => req.destroy(new Error("Tool relay timed out")));
    req.end(JSON.stringify({ session, ...body as object }));
  });
}
const server = new Server({ name: "agent-gateway-client-tools", version: "0.1.0" }, { capabilities: { tools: {} }, instructions: gatewayToolInstructions(process.env.AGENT_GATEWAY_PROVIDER, process.env.AGENT_GATEWAY_OFFICE_JS === "1") });
server.setRequestHandler(ListToolsRequestSchema, async () => call("/internal/tools", {}));
server.setRequestHandler(CallToolRequestSchema, async request => call("/internal/tool", { name: request.params.name, arguments: request.params.arguments ?? {} }));
await server.connect(new StdioServerTransport());
