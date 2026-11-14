import fs from "node:fs";
import path from "node:path";
import { z } from "zod";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { AjvJsonSchemaValidator } from "@modelcontextprotocol/sdk/validation/ajv";

export const McpServerConfig = z.object({
  id: z.string().regex(/^[a-z0-9][a-z0-9_-]{0,47}$/),
  name: z.string().trim().min(1).max(100),
  transport: z.enum(["stdio", "http"]),
  target: z.string().trim().min(1).max(2000),
  args: z.array(z.string().max(2000)).max(100).default([]),
}).strict().superRefine((value, ctx) => {
  if (value.transport !== "http") return;
  try {
    const url = new URL(value.target);
    if (!["https:", "http:"].includes(url.protocol) || url.username || url.password) throw new Error();
  } catch { ctx.addIssue({ code: "custom", message: "Use an HTTP(S) URL without embedded credentials", path: ["target"] }); }
});
export type McpServerConfig = z.infer<typeof McpServerConfig>;

/** Bridge-owned connections. Configuration is shared, execution state is isolated by pane. */
export class McpServers {
  private configs: McpServerConfig[];
  private clients = new Map<string, Promise<Client>>();
  constructor(private readonly file: string, private readonly options: { callTimeoutMs?: number } = {}) {
    try { this.configs = z.array(McpServerConfig).max(30).parse(JSON.parse(fs.readFileSync(file, "utf8"))); }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw new Error("Invalid MCP server configuration"); this.configs = []; }
  }
  async reload() {
    let next: McpServerConfig[];
    try { next = z.array(McpServerConfig).max(30).parse(JSON.parse(fs.readFileSync(this.file, "utf8"))); }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw new Error("Invalid MCP server configuration"); next = []; }
    if (JSON.stringify(next) !== JSON.stringify(this.configs)) { await this.close(); this.configs = next; }
  }
  list() { return this.configs.map(item => ({ ...item, args: [...item.args] })); }
  private persist(next: McpServerConfig[]) {
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    const temporary = `${this.file}.tmp`;
    fs.writeFileSync(temporary, JSON.stringify(next, null, 2), { mode: 0o600 });
    fs.renameSync(temporary, this.file);
    this.configs = next;
  }
  async save(input: unknown) {
    const config = McpServerConfig.parse(input);
    const next = [...this.configs.filter(item => item.id !== config.id), config];
    if (next.length > 30) throw new Error("Maximum 30 MCP servers");
    this.persist(next);
    await this.closeServer(config.id);
  }
  async remove(id: string) { this.persist(this.configs.filter(item => item.id !== id)); await this.closeServer(id); }
  private async closeServer(id: string) {
    for (const key of [...this.clients.keys()]) if (key.endsWith(`:${id}`)) await this.closeKey(key);
  }
  private async closeKey(key: string) {
    const pending = this.clients.get(key); this.clients.delete(key);
    if (pending) await pending.then(client => client.close(), () => undefined).catch(() => undefined);
  }
  async closeSessionServer(session: string, id: string) { await this.closeKey(`${session}:${id}`); }
  async closeSession(session: string) {
    await Promise.all([...this.clients.keys()].filter(key => key.startsWith(`${session}:`)).map(key => this.closeKey(key)));
  }
  async close() { await Promise.all([...this.clients.keys()].map(key => this.closeKey(key))); }
  private client(session: string, id: string, signal?: AbortSignal): Promise<Client> {
    signal?.throwIfAborted();
    const config = this.configs.find(item => item.id === id);
    if (!config) throw new Error("MCP server not found");
    const key = `${session}:${id}`;
    const existing = this.clients.get(key); if (existing) return existing;
    const client = new Client({ name: "office-partner", version: "0.1.21" });
    // Do not give child processes bridge credentials or provider tokens from process.env.
    const env = Object.fromEntries(["PATH", "Path", "HOME", "USERPROFILE", "LOCALAPPDATA", "APPDATA", "SYSTEMROOT", "SystemRoot", "TEMP", "TMP", "COMSPEC", "PATHEXT"].flatMap(name => process.env[name] ? [[name, process.env[name]!]] : []));
    const transport = config.transport === "stdio"
      ? new StdioClientTransport({ command: config.target, args: config.args, env, stderr: "ignore" })
      : new StreamableHTTPClientTransport(new URL(config.target));
    const pending = client.connect(transport, { timeout: 15_000, signal }).then(() => client).catch(async error => {
      if (this.clients.get(key) === pending) this.clients.delete(key);
      await client.close().catch(() => undefined); throw error;
    });
    this.clients.set(key, pending);
    client.onclose = () => { if (this.clients.get(key) === pending) this.clients.delete(key); };
    return pending;
  }
  async tools(session: string, id: string, signal?: AbortSignal) {
    return this.listTools(await this.client(session, id, signal), signal);
  }
  private async listTools(client: Client, signal?: AbortSignal) {
    const tools = []; const cursors = new Set<string>(); let cursor: string | undefined;
    do {
      const page = await client.listTools(cursor ? { cursor } : {}, { timeout: 15_000, signal });
      tools.push(...page.tools);
      if (tools.length > 500) throw new Error("MCP catalog exceeds 500 tools");
      cursor = page.nextCursor;
      if (cursor && cursors.has(cursor)) throw new Error("MCP server repeated a pagination cursor");
      if (cursor) cursors.add(cursor);
    } while (cursor);
    if (JSON.stringify(tools).length > 1_000_000) throw new Error("MCP catalog is too large");
    return tools;
  }
  async call(session: string, id: string, name: string, args: Record<string, unknown>, signal?: AbortSignal) {
    const client = await this.client(session, id, signal);
    const tool = (await this.listTools(client, signal)).find(item => item.name === name);
    if (!tool) throw new Error("MCP tool not found");
    const checked = new AjvJsonSchemaValidator().getValidator(tool.inputSchema)(args);
    if (!checked.valid) throw new Error(`Invalid MCP arguments: ${checked.errorMessage}`);
    signal?.throwIfAborted();
    const result = await client.callTool({ name, arguments: args }, undefined, { timeout: this.options.callTimeoutMs ?? 60_000, signal });
    if (JSON.stringify(result).length > 4_000_000) throw new Error("MCP result exceeds 4 MB; request a smaller result. The operation may already have completed.");
    return result;
  }
}
