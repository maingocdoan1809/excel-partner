/**
 * Drive a REAL provider CLI end to end against an in-memory workbook.
 *
 *   npm run build:server
 *   npx tsx scripts/live-smoke.ts <claude|copilot|codex> ["prompt"] [--tls] [--cancel-after=ms] [--turns=N]
 *
 * Starts the bridge, connects a fake task pane, sends the prompt(s), prints the
 * normalised event stream and the resulting workbook state.
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { PaneClient, getFreePort } from "../tests/helpers.js";
import { FakeWorkbook } from "../tests/support/fake-workbook.js";
import type { ProviderId } from "../protocol/src/index.js";
import { FakeAnthropic } from "../tests/support/fake-anthropic.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const flags = new Map<string, string>(
  args
    .filter((a) => a.startsWith("--"))
    .map((a) => {
      const [key = "", value = ""] = a.slice(2).split("=");
      return [key, value] as [string, string];
    }),
);
const positional = args.filter((a) => !a.startsWith("--"));
const provider = (positional[0] ?? "codex") as ProviderId;
const prompts = positional.slice(1);
if (prompts.length === 0) {
  prompts.push(
    "Look at my selected row. Add a column D header 'Total' in D1 and put a formula =B2*C2 in D2 and =B3*C3 in D3, then bold the header row.",
  );
}

// --fake-claude: run the real `claude` binary against a scripted fake Anthropic API.
let fake: FakeAnthropic | null = null;
if (flags.has("fake-claude")) {
  fake = new FakeAnthropic([
    { tool: "get_selection", input: {} },
    { tool: "set_formula", input: { address: "D2", formula: "B2*C2" } },
    { tool: "format_range", input: { address: "A1:D1", format: { bold: true } } },
    { text: "Done: added the formula in D2 and bolded the header row." },
  ]);
  await fake.start();
  process.env.ANTHROPIC_BASE_URL = fake.url;
  process.env.ANTHROPIC_API_KEY = "sk-ant-fake";
}

const useTls = flags.has("tls");
const port = await getFreePort();
const config = loadConfig({
  host: "127.0.0.1",
  port,
  logLevel: (process.env.BRIDGE_LOG_LEVEL as any) ?? "warn",
  ...(useTls ? { publicBaseUrl: `https://127.0.0.1:${port}` } : { tls: null, publicBaseUrl: `http://127.0.0.1:${port}` }),
  mcp: { command: process.execPath, args: [path.join(root, "dist/excel_mcp/src/cli.js")] },
});
const bridge = new Bridge(config);
const { url } = await bridge.listen();
const wb = new FakeWorkbook();

import fs from "node:fs";
const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`, {
  ca: useTls ? fs.readFileSync(path.join(root, "certs/localhost.crt")) : undefined,
});
pane.onMessage = async (m) => {
  if (m.type === "tool_request") {
    const result = await wb.handle(m.tool, m.args);
    console.log(`  [pane] ${m.tool} ${JSON.stringify(m.args)} -> ${result.ok ? "ok" : "ERR " + JSON.stringify(result.error)}`);
    pane.send({ type: "tool_result", callId: m.callId, result });
  } else if (m.type === "message_delta") {
    process.stdout.write(m.text);
  } else {
    console.log(`\n  [${m.type}] ${JSON.stringify(m).slice(0, 220)}`);
  }
};
pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef" });
await pane.waitFor((m) => m.type === "hello_ok");
pane.send({ type: "set_provider", provider });
const status = await pane.waitFor((m) => m.type === "provider_status", 20_000);
console.log("provider_status:", JSON.stringify((status as any).detection));

let n = 0;
for (const prompt of prompts) {
  const requestId = `smoke-${++n}`;
  console.log(`\n=== turn ${n}: ${prompt}`);
  const seenBefore = pane.received.length;
  pane.send({ type: "prompt", requestId, prompt, selection: null });
  const cancelAfter = flags.get("cancel-after");
  if (cancelAfter) setTimeout(() => pane.send({ type: "cancel", requestId }), Number(cancelAfter));
  await pane.waitFor(
    (m, i) => i >= seenBefore && m.type === "turn_finished" && (m as any).requestId === requestId,
    Number(flags.get("timeout") ?? 180_000),
  );
}

console.log("\n=== workbook state");
for (const [name, cells] of wb.sheets) console.log(name, JSON.stringify(Object.fromEntries(cells)));
console.log("formats:", JSON.stringify(wb.formats));
console.log("calls:", wb.calls.map((c) => `${c.tool}${c.ok ? "" : "(err)"}`).join(", "));

await pane.close();
await bridge.close();
process.exit(0);
