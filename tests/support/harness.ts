import path from "node:path";
import { fileURLToPath } from "node:url";
import { Bridge } from "../../bridge/src/server.js";
import { loadConfig } from "../../bridge/src/config.js";
import type { ProviderId, ServerMessage } from "../../protocol/src/index.js";
import { PaneClient, getFreePort } from "../helpers.js";
import { FakeWorkbook } from "./fake-workbook.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

export interface TurnResult {
  events: ServerMessage[];
  finished: Extract<ServerMessage, { type: "turn_finished" }>;
  text: string;
}

/**
 * Bridge + fake task pane + in-memory workbook, wired the way production is:
 * the provider CLI launches the real `dist/excel_mcp` stdio server, which
 * calls back into the bridge, which routes to this pane over WebSocket.
 * Requires `npm run build:server` first.
 */
export class Harness {
  readonly wb = new FakeWorkbook();
  private seq = 0;

  private constructor(
    readonly bridge: Bridge,
    readonly pane: PaneClient,
    readonly url: string,
  ) {
    pane.onMessage = (m) => {
      if (m.type === "tool_request") {
        void this.wb.handle(m.tool, m.args).then((result) => pane.send({ type: "tool_result", callId: m.callId, result }));
      }
    };
  }

  static async start(provider: ProviderId, overrides: Parameters<typeof loadConfig>[0] = {}): Promise<Harness> {
    const port = await getFreePort();
    const config = loadConfig({
      host: "127.0.0.1",
      port,
      tls: null,
      publicBaseUrl: `http://127.0.0.1:${port}`,
      logLevel: (process.env.HARNESS_LOG_LEVEL as "debug" | "error" | undefined) ?? "error",
      mcp: { command: process.execPath, args: [path.join(root, "dist/excel_mcp/src/cli.js")] },
      ...overrides,
    });
    const bridge = new Bridge(config, { usageFile: null });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    const harness = new Harness(bridge, pane, url);
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef" });
    await pane.waitFor((m) => m.type === "hello_ok");
    pane.send({ type: "set_provider", provider });
    await pane.waitFor((m) => m.type === "provider_status", 30_000);
    return harness;
  }

  async turn(prompt: string, options: { cancelAfterMs?: number; timeoutMs?: number } = {}): Promise<TurnResult> {
    const requestId = `h-${++this.seq}`;
    const from = this.pane.received.length;
    this.pane.send({ type: "prompt", requestId, prompt, selection: null });
    if (options.cancelAfterMs !== undefined) {
      setTimeout(() => this.pane.send({ type: "cancel", requestId }), options.cancelAfterMs);
    }
    const finished = (await this.pane.waitFor(
      (m, i) => i >= from && m.type === "turn_finished" && m.requestId === requestId,
      options.timeoutMs ?? 120_000,
    )) as TurnResult["finished"];
    // Let any trailing events (a duplicate finish would show up here) arrive.
    await new Promise((r) => setTimeout(r, 150));
    const events = this.pane.received.slice(from).filter((m) => "requestId" in m && m.requestId === requestId);
    const text = events.map((e) => (e.type === "message_delta" ? e.text : "")).join("");
    return { events, finished, text };
  }

  async close(): Promise<void> {
    await this.pane.close();
    await this.bridge.close();
  }
}
