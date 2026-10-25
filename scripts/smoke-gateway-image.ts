import assert from "node:assert/strict";
import { GatewayServer } from "../gateway/src/server.js";
import { ConfigSchema } from "../gateway/src/config.js";
import { createDefaultRegistry } from "../providers/src/registry.js";
import type { ProviderId } from "../protocol/src/index.js";
import type { ApiMessage, ChatMessage } from "../gateway/src/protocol.js";
import { gatewayProbeImage } from "../tests/support/gateway-probe-image.js";

const provider = (process.argv[2] ?? "copilot") as ProviderId;
const nativeModel = process.argv[3] ?? "default";
const token = "gateway-synthetic-image-probe-only";
const config = { ...ConfigSchema.parse({ warmupOnStart: false, enabledProviders: [provider], defaultRoute: {
  provider, model: nativeModel === "default" ? null : nativeModel,
} }), port: 0 };
const registry = createDefaultRegistry({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
const adapter = registry.require(provider);
const start = adapter.startSession.bind(adapter);
let starts = 0;
adapter.startSession = async ctx => { starts++; return start(ctx); };
const server = new GatewayServer({ config, token, tls: null, registry });
const url = await server.listen();
const messages: ChatMessage[] = [];
const tools = [{ name: "image_probe", description: "Return the current image from the client for visual inspection.",
  input_schema: { type: "object", properties: {}, additionalProperties: false } }];
const post = async () => {
  const response = await fetch(`${url}/v1/messages`, { method: "POST", headers: { "x-api-key": token, "content-type": "application/json" },
    body: JSON.stringify({ model: `${provider}/${nativeModel}`, max_tokens: 256, messages, tools }) });
  const body = await response.json() as ApiMessage;
  assert.equal(response.status, 200, JSON.stringify(body));
  return body;
};
try {
  for (const reverse of [false, true]) {
    messages.push({ role: "user", content: "Call image_probe once to inspect a NEW image. Identify the dominant color in its LEFT half and RIGHT half. Reply exactly LEFT=<color> RIGHT=<color>. Use this new image, not any previous image or metadata. Do not call other tools." });
    const call = await post();
    assert.equal(call.stop_reason, "tool_use");
    const blocks = call.content.filter(block => block.type === "tool_use");
    assert.equal(blocks.length, 1);
    assert.equal(blocks[0]!.name, "image_probe");
    messages.push({ role: "assistant", content: call.content }, { role: "user", content: [{ type: "tool_result", tool_use_id: blocks[0]!.id,
      content: [{ type: "image", source: { type: "base64", media_type: "image/png", data: gatewayProbeImage(reverse, 512, 256).toString("base64") } }] }] });
    const answer = await post();
    assert.equal(answer.stop_reason, "end_turn");
    const text = answer.content.filter(block => block.type === "text").map(block => String(block.text)).join(" ");
    assert.match(text, reverse ? /LEFT\s*=\s*BLUE\b.*RIGHT\s*=\s*RED\b/is : /LEFT\s*=\s*RED\b.*RIGHT\s*=\s*BLUE\b/is);
    messages.push({ role: "assistant", content: answer.content });
    console.log(`${provider}: ${text}`);
  }
  assert.equal(starts, 1, "Image turns must reuse the same worker");
  console.log(`PASS ${provider}: real CLI saw two different tool-result images, correct colors, one worker.`);
} finally { await server.close(); }
