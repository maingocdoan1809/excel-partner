import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { normalizeCommands } from "../providers/src/commands.js";
import type { SendTurnInput } from "../providers/src/types.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | undefined;
let pane: PaneClient | undefined;
afterEach(async () => { await pane?.close(); await bridge?.close(); });

async function setup() {
  const inputs: SendTurnInput[] = [];
  class Adapter extends MockAdapter {
    async listCommands() { return normalizeCommands([{ name: "skills", description: "Skills" }, { name: "vim", description: "Vim", execution: "unsupported", reason: "Terminal only" }]); }
    override async *sendTurn(input: SendTurnInput) { inputs.push(input); yield* super.sendTurn(input); }
  }
  const port = await getFreePort();
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }),
    { usageFile: null, registry: new ProviderRegistry().register(new Adapter({ id: "codex", reply: "Listed skills." })) });
  const { url } = await bridge.listen();
  pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "commands-workbook" });
  const hello = await pane.waitFor(message => message.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("No hello");
  pane.send({ type: "set_provider", provider: "codex" });
  await pane.waitFor(message => message.type === "provider_status");
  return { url, inputs, headers: { authorization: `Bearer ${hello.sessionToken}` } };
}

describe("provider commands through the bridge", () => {
  it("authenticates discovery and merges provider commands with add-in controls", async () => {
    const { url, headers } = await setup();
    expect((await fetch(`${url}/session/commands?provider=codex`)).status).toBe(401);
    expect((await fetch(`${url}/session/commands?provider=unknown`, { headers })).status).toBe(400);
    const result = await (await fetch(`${url}/session/commands?provider=codex`, { headers })).json();
    expect(result.commands).toEqual(expect.arrayContaining([
      expect.objectContaining({ name: "skills", execution: "native" }), expect.objectContaining({ name: "model", execution: "client" }),
    ]));
  });
  it("dispatches exact native arguments without replay or saved-reference context", async () => {
    const { inputs } = await setup();
    pane!.send({ type: "prompt", requestId: "command", prompt: "/skills list", command: true, selection: null, context: "REPLAY MUST NOT REACH COMMAND" });
    await pane!.waitFor(message => message.type === "turn_finished" && message.requestId === "command");
    expect(inputs).toHaveLength(1);
    expect(inputs[0]).toMatchObject({ prompt: "/skills list", command: { name: "skills", args: "list" } });
    expect(inputs[0]!.prompt).not.toContain("REPLAY");
  });
  it.each(["/vim", "/not-a-command"])("rejects %s without invoking the model", async prompt => {
    const { inputs } = await setup();
    pane!.send({ type: "prompt", requestId: "bad", prompt, command: true, selection: null });
    expect(await pane!.waitFor(message => message.type === "turn_finished" && message.requestId === "bad")).toMatchObject({ reason: "error" });
    expect(inputs).toHaveLength(0);
  });
});
