// Run after npm run build:server. Uses real CLIs, but never submits a model task.
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Harness } from "../dist/tests/support/harness.js";

const cases = { copilot: ["/skills list", /skills/i], claude: ["/context", /context|tokens/i],
  codex: ["/skills", /skills/i], opencode: ["/skills", /skills/i] };
const selected = process.argv[2] ? [process.argv[2]] : Object.keys(cases);
for (const provider of selected) {
  if (!cases[provider]) throw new Error(`Unknown provider: ${provider}`);
  const previousHome = process.env.EXCEL_AI_HOME;
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "excel-command-smoke-"));
  process.env.EXCEL_AI_HOME = directory;
  let harness;
  try {
    harness = await Harness.start(provider);
    const hello = harness.pane.received.find(message => message.type === "hello_ok");
    assert(hello);
    const response = await fetch(`${harness.url}/session/commands?provider=${provider}`, { headers: { authorization: `Bearer ${hello.sessionToken}` } });
    const catalog = await response.json();
    assert.equal(response.status, 200, JSON.stringify(catalog));
    const [prompt, expected] = cases[provider];
    assert(catalog.commands.some(command => command.name === prompt.split(" ")[0].slice(1) && command.execution === "native"));
    harness.pane.send({ type: "prompt", requestId: "command-smoke", prompt, command: true, selection: null, context: "Do not turn this into a model question." });
    const finished = await harness.pane.waitFor(message => message.type === "turn_finished" && message.requestId === "command-smoke", 45_000);
    assert.equal(finished.reason, "completed", JSON.stringify(harness.pane.received.filter(message => message.type === "error")));
    const text = harness.pane.received.flatMap(message => message.type === "message_completed" ? [message.text] : []).join("\n");
    assert.match(text, expected);
    console.log(`${provider}: ${catalog.commands.length} commands; ${prompt} passed without a model task.`);
  } finally {
    await harness?.close();
    if (previousHome === undefined) delete process.env.EXCEL_AI_HOME; else process.env.EXCEL_AI_HOME = previousHome;
    await fs.promises.rm(directory, { recursive: true, force: true, maxRetries: 20, retryDelay: 100 });
  }
}
