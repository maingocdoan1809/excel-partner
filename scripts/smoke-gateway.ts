import { executeOfficeJs } from "../addin/taskpane/executor/office-js.js";
import { installOfficeMock, MockWorkbook, OfficeError } from "../tests/support/office-mock.js";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { createDefaultRegistry } from "../providers/src/registry.js";
import { GatewayServer } from "../gateway/src/server.js";
import { ConfigSchema } from "../gateway/src/config.js";
import type { ChatMessage, ApiMessage } from "../gateway/src/protocol.js";
import type { ProviderId } from "../protocol/src/index.js";
import { gatewayProbeImage } from "../tests/support/gateway-probe-image.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const provider = (process.argv[2] ?? "copilot") as ProviderId;
const nativeModel = process.argv[3] && !process.argv[3]!.startsWith("--") ? process.argv[3]! : "default";
const office = process.argv.includes("--office");
const line = process.argv.includes("--line");
const redactImages = process.argv.includes("--redact-images");
const expireWorker = process.argv.includes("--expire-worker");
const restartGateway = process.argv.includes("--restart-gateway");
const useConversationId = restartGateway || process.argv.includes("--conversation-id");
const metadata = useConversationId ? { user_id: JSON.stringify({ session_id: randomUUID(), user: "synthetic-probe", device_id: "probe" }) } : undefined;
const probeDirectory = restartGateway ? fs.mkdtempSync(path.join(os.tmpdir(), "gateway-restart-probe-")) : undefined;
const collision = process.argv.includes("--collision") || office;
const readTool = collision ? "read_surface_context" : "read_probe";
const writeTool = office ? "execute_office_js" : collision ? "execute" : "write_probe";
const token = "gateway-smoke-test-token-not-a-real-credential";
const config = { ...ConfigSchema.parse({ warmupOnStart: false, enabledProviders: [provider], defaultRoute: { provider, model: nativeModel === "default" ? null : nativeModel } }), port: 0 };
if (expireWorker) config.idleTimeoutMs = 80;
const registry = createDefaultRegistry({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
const adapter = registry.require(provider);
const originalStart = adapter.startSession.bind(adapter);
const originalTurn = adapter.sendTurn.bind(adapter);
let starts = 0;
const runtimes: string[] = [];
const prompts: string[] = [];
const resumes: Array<string | undefined> = [];
adapter.startSession = async ctx => { starts++; resumes.push(ctx.resumeSessionId); return originalStart(ctx); };
adapter.sendTurn = input => { runtimes.push(input.providerSessionId); prompts.push(input.prompt); return originalTurn(input); };
const makeServer = () => new GatewayServer({ config, token, tls: null, registry,
  ...(probeDirectory ? { conversationFile: path.join(probeDirectory, "conversations.json") } : {}),
  log: event => console.log(JSON.stringify(event)) });
let server = makeServer();
let url = await server.listen();
const sentinel = `ZEBRA-${randomUUID()}`;
const writes: string[] = [];
const book = new MockWorkbook();
book.addSheet("Sheet1");
book.addSheet("Sheet3");
if (collision) {
  const addSheet = book.addSheet.bind(book);
  book.addSheet = name => {
    if (book.sheets.some(sheet => sheet.name.toLowerCase() === name.toLowerCase())) throw new OfficeError("ItemAlreadyExists", "A sheet with this name already exists");
    return addSheet(name);
  };
  installOfficeMock(book);
}
const messages: ChatMessage[] = [{ role: "user", content: collision
  ? `Call read_surface_context, then use the remote ${writeTool} tool to ensure Sheet3 exists and write the returned value verbatim into cell A1. Pass the async function BODY: context is already in scope. Do not define main or wrap in Excel.run. Read back A1 after writing. ${line ? "Also create a straight line shape named QA_LINE on Sheet3 from point (200,50) to point (300,150), then read back its name and geometry." : ""} Do not run Office.js locally. Reply PROBE_DONE after the client confirms the actual cell value.`
  : "Call read_probe, then call write_probe with the exact value read_probe returned. Do not invent the value. Then reply PROBE_DONE. Use only the registered tools." }];
const tools = [
  { name: readTool, description: "Read the test value and workbook context from the remote client", input_schema: { type: "object", properties: {}, additionalProperties: false } },
  { name: writeTool, description: collision ? "Execute Office.js code in the Excel client. code is the BODY of an async function with context already in scope, executed inside Excel.run. Use statements and return plain data directly, not a main function declaration." : "Write the exact test value",
    input_schema: collision ? { type: "object", properties: { code: { type: "string" }, explanation: { type: "string" } }, required: ["code"], additionalProperties: false }
      : { type: "object", properties: { value: { type: "string" } }, required: ["value"], additionalProperties: false } },
];
try {
  for (let i = 0; i < 10; i++) {
    const res = await fetch(`${url}/v1/messages`, { method: "POST", headers: { "x-api-key": token, "content-type": "application/json" },
      body: JSON.stringify({ model: `${provider}/${nativeModel}`, max_tokens: 2048, messages, tools, metadata }) });
    const result = await res.json() as ApiMessage;
    if (!res.ok) throw new Error(JSON.stringify(result));
    console.log(JSON.stringify({ stop: result.stop_reason, content: result.content }));
    messages.push({ role: "assistant", content: result.content });
    if (result.stop_reason === "end_turn") {
      assert.equal(writes.length, 1); assert.equal(writes[0], sentinel);
      if (collision) { assert.equal(book.get("Sheet3", "A1"), sentinel); assert.equal(book.sheets.length, 2); }
      if (line) {
        const shape = book.sheet("Sheet3").shapeList.find(shape => shape.name === "QA_LINE");
        assert.ok(shape, "The model must actually create the line through the client tool");
        assert.equal(shape.type, "Line");
        assert.deepEqual([shape.left, shape.top, shape.width, shape.height], [200, 50, 100, 100]);
        assert.equal(shape.lineProps.connectorType, "Straight");
      }
      assert.match(result.content.filter(b => b.type === "text").map(b => b.text).join(""), /PROBE_DONE/);
      if (redactImages) {
        for (const message of messages) for (const block of Array.isArray(message.content) ? message.content : []) {
          if (block.type === "tool_result" && Array.isArray(block.content)) block.content = block.content.map(part =>
            part.type === "image" ? { type: "text", text: "[image previously shown — redacted from history]" } : part);
        }
      }
      messages.push({ role: "user", content: "Without calling any tools, reply MEMORY_CHECK followed by the exact value you read in the previous turn." });
      if (expireWorker) await new Promise(resolve => setTimeout(resolve, 150));
      if (restartGateway) { await server.close(); registry.register(adapter); server = makeServer(); url = await server.listen(); }
      const followup = await fetch(`${url}/v1/messages`, { method: "POST", headers: { "x-api-key": token, "content-type": "application/json" },
        body: JSON.stringify({ model: `${provider}/${nativeModel}`, max_tokens: 2048, messages, tools, metadata }) });
      const remembered = await followup.json() as ApiMessage;
      assert.equal(followup.status, 200, JSON.stringify(remembered));
      assert.equal(remembered.stop_reason, "end_turn");
      assert.ok(remembered.content.some(b => b.type === "text" && String(b.text).includes(sentinel)), JSON.stringify(remembered));
      assert.equal(starts, expireWorker || restartGateway ? 2 : 1, "Follow-ups must reuse the live worker or resume its confirmed native session");
      if (expireWorker || restartGateway) assert.ok(resumes[1], "Recovery must supply the native resume ID");
      else assert.equal(new Set(runtimes).size, 1);
      assert.equal(prompts.length, 2);
      assert.ok(!prompts[1]!.includes(sentinel), "Old history must not be replayed to the live worker");
      console.log(`PASS ${provider}${collision ? " (remote execute collision)" : ""}${redactImages ? " (Office image redaction)" : ""}: real CLI tools, two completed turns, ${restartGateway ? "persisted native resume after gateway restart" : expireWorker ? "native resume after idle expiry" : "one worker"}, remembered value without history replay.`);
      break;
    }
    const results = [];
    for (const block of result.content.filter(b => b.type === "tool_use")) {
      const input = block.input as { value: string; code: string };
      let value: unknown;
      let failed = false;
      if (block.name === writeTool) {
        if (collision) {
          const before = book.get("Sheet3", "A1");
          try {
            const executed = await executeOfficeJs({ code: input.code }) as { result: unknown };
            value = executed.result === null ? { success: true } : { success: true, result: executed.result };
          } catch (error) {
            failed = true;
            value = { success: false, error: error instanceof Error ? error.message : String(error) };
          } finally {
            const after = book.get("Sheet3", "A1");
            if (after !== before) { assert.equal(after, sentinel); writes.push(String(after)); }
          }
        } else { writes.push(input.value); value = { written: true }; }
      } else {
        assert.equal(block.name, readTool);
        const data = { value: sentinel, sheets: book.sheets.map(sheet => sheet.name) };
        value = collision ? { success: true, result: data } : data;
      }
      const content = redactImages && block.name === readTool ? [
        { type: "text", text: JSON.stringify(value) },
        { type: "image", source: { type: "base64", media_type: "image/png", data: gatewayProbeImage().toString("base64") } },
      ] : JSON.stringify(value);
      results.push({ type: "tool_result", tool_use_id: block.id, content, ...(failed ? { is_error: true } : {}) });
    }
    assert.ok(results.length);
    messages.push({ role: "user", content: results });
    if (i === 9) throw new Error("Too many tool turns");
  }
} finally { await server.close(); if (probeDirectory) fs.rmSync(probeDirectory, { recursive: true, force: true }); }
