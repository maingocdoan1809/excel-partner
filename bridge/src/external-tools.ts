import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

type Service = "browser";

/** External MCP connections are isolated per workbook session and created only on first use. */
export class ExternalTools {
  private readonly clients = new Map<string, Promise<Client>>();
  private readonly outputDirs = new Map<string, string>();

  private cleanOutput(key: string, dir?: string): void {
    if (!dir) return;
    if (this.outputDirs.get(key) === dir) this.outputDirs.delete(key);
    fs.rmSync(dir, { recursive: true, force: true });
  }

  private client(sessionId: string, service: Service, signal?: AbortSignal): Promise<Client> {
    const key = `${sessionId}:${service}`;
    const existing = this.clients.get(key);
    if (existing) return existing;
    const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-partner-browser-"));
    if (outputDir) this.outputDirs.set(key, outputDir);
    const client = new Client({ name: "excel-partner", version: "0.1.0" });
    const pending: Promise<Client> = Promise.resolve().then(async () => {
      try {
        const transport = new StdioClientTransport({
          command: "npx",
          args: ["--yes", "@playwright/mcp@latest", "--headless", "--isolated",
            "--output-dir", outputDir,
            ...(process.env.BRIDGE_PLAYWRIGHT_EXECUTABLE ? ["--executable-path", process.env.BRIDGE_PLAYWRIGHT_EXECUTABLE] : []),
          ],
          cwd: outputDir,
          stderr: "ignore",
        });
        await client.connect(transport, { timeout: 60_000, signal });
        signal?.throwIfAborted();
        return client;
      } catch (error) {
        if (this.clients.get(key) === pending) this.clients.delete(key);
        await client.close().catch(() => undefined);
        this.cleanOutput(key, outputDir);
        throw error;
      }
    });
    this.clients.set(key, pending);
    client.onclose = () => {
      if (this.clients.get(key) === pending) this.clients.delete(key);
      this.cleanOutput(key, outputDir);
    };
    return pending;
  }

  async browser(sessionId: string, tool: string, args: Record<string, unknown>, signal?: AbortSignal): Promise<unknown> {
    const cancel = () => { void this.closeSession(sessionId, "browser"); };
    signal?.addEventListener("abort", cancel, { once: true });
    try {
      signal?.throwIfAborted();
      const client = await this.client(sessionId, "browser", signal);
      const tools = [];
      let cursor: string | undefined;
      do {
        const page = await client.listTools(cursor ? { cursor } : undefined, { timeout: 15_000, signal });
        tools.push(...page.tools);
        cursor = page.nextCursor;
      } while (cursor);
      if (tool === "list_tools") return { tools };
      if (!tools.some(candidate => candidate.name === tool)) throw new Error(`Unknown Playwright tool '${tool}'. Call list_tools first.`);
      return await client.callTool({ name: tool, arguments: args }, undefined, { timeout: 60_000, signal });
    } finally { signal?.removeEventListener("abort", cancel); }
  }

  async closeSession(sessionId: string, service?: Service): Promise<void> {
    const pending = [...this.clients].filter(([key]) => service ? key === `${sessionId}:${service}` : key.startsWith(`${sessionId}:`))
      .map(([key, connection]) => ({ key, connection, outputDir: this.outputDirs.get(key) }));
    for (const { key } of pending) this.clients.delete(key);
    await Promise.allSettled(pending.map(async ({ connection }) => (await connection).close()));
    for (const { key, outputDir } of pending) this.cleanOutput(key, outputDir);
  }

  async close(): Promise<void> {
    const sessions = new Set([...this.clients.keys()].map(key => key.slice(0, key.lastIndexOf(":"))));
    await Promise.allSettled([...sessions].map(sessionId => this.closeSession(sessionId)));
  }
}
