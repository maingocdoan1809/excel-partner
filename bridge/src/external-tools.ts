import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { BROWSER_FRAME_MAX_BASE64 } from "../../protocol/src/messages.js";
import { FRAME_AFTER, FRAME_ON_FIRST_READ, describeBrowserAction, extractImage, parsePageInfo, type ToolOutput } from "./browser-frames.js";
import { BrowserGuard, GUARDED_TOOLS } from "./browser-guard.js";
import { SIGNALS_FUNCTION, buildDiagnostics, modalState, parseSignals } from "./browser-diagnostics.js";

type Service = "browser";

/** One picture of the assistant's browser, ready to be sent to the task pane. */
export interface BrowserFrame { seq: number; label: string; url?: string; title?: string; mimeType: "image/jpeg" | "image/png"; data: string }
export type FrameSink = (frame: BrowserFrame) => void;

interface FrameState { seq: number; token: number; shown: boolean; url?: string; title?: string }

/** A browser limited to the sites the user approved. It starts signed out, or with a session the user saved with `excel-partner login` (the file is copied, never modified). */
export interface BrowserProfile { allowedOrigins: string[]; visible?: boolean; storageState?: string }

function bypassEntry(origin: string): string | null {
  try {
    const url = new URL(origin);
    return url.port ? `${url.hostname}:${url.port}` : url.hostname;
  } catch { return null; }
}

/** External MCP connections are isolated per workbook session and created only on first use. */
export class ExternalTools {
  private readonly clients = new Map<string, Promise<Client>>();
  private readonly outputDirs = new Map<string, string>();
  private readonly profiles = new Map<string, BrowserProfile>();
  /** Calls to one session's browser run one at a time: a screenshot taken during a navigation fails. */
  private readonly queues = new Map<string, Promise<unknown>>();
  private readonly frames = new Map<string, FrameState>();
  private readonly guard = new BrowserGuard();

  /** Next browser for this session is limited to these origins; any running one is closed. */
  async useProfile(sessionId: string, profile: BrowserProfile | null): Promise<void> {
    if (profile) this.profiles.set(sessionId, profile); else this.profiles.delete(sessionId);
    await this.closeSession(sessionId, "browser");
  }

  hasProfile(sessionId: string): boolean { return this.profiles.has(sessionId); }

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
    const profile = this.profiles.get(sessionId);
    const profileArgs: string[] = [];
    if (profile) {
      if (profile.storageState) {
        const copy = path.join(outputDir, "state.json");
        fs.copyFileSync(profile.storageState, copy);
        try { fs.chmodSync(copy, 0o600); } catch { /* best effort */ }
        profileArgs.push("--storage-state", copy);
      }
      // Allowed origins at the page level, plus a dead proxy for everything else so IP addresses are blocked too.
      profileArgs.push("--allowed-origins", profile.allowedOrigins.join(";"),
        "--proxy-server", "http://127.0.0.1:9", "--proxy-bypass", profile.allowedOrigins.map(bypassEntry).filter(Boolean).join(","));
    }
    const client = new Client({ name: "excel-partner", version: "0.1.0" });
    const pending: Promise<Client> = Promise.resolve().then(async () => {
      try {
        const transport = new StdioClientTransport({
          command: "npx",
          args: ["--yes", "@playwright/mcp@latest", ...(profile?.visible ? [] : ["--headless"]), "--isolated",
            "--output-dir", outputDir, ...profileArgs,
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

  /**
   * Run one Playwright MCP tool. With `onFrame`, a picture of the page is sent after actions that change it, so the
   * person can watch. The picture is taken after the result has been handed back and never delays or alters it.
   */
  browser(sessionId: string, tool: string, args: Record<string, unknown>, signal?: AbortSignal, onFrame?: FrameSink, turn?: string): Promise<unknown> {
    const run = () => this.browserCall(sessionId, tool, args, signal, onFrame, turn);
    const result = (this.queues.get(sessionId) ?? Promise.resolve()).then(run, run);
    this.queues.set(sessionId, result.catch(() => undefined));
    return result;
  }

  /**
   * When an action fails, the assistant gets the page as it is now (a picture and the dialogs, invalid fields and
   * messages on it) together with the error, and an identical action that keeps failing is refused. Without this a
   * model that cannot see the screen tends to repeat the same click until it is stopped.
   */
  private async explain(client: Client, tool: string, args: Record<string, unknown>, base: ToolOutput | null,
    info: { kind: "failed" | "blocked"; attempts: number; reason?: "same" | "streak" }): Promise<ToolOutput> {
    const modal = base ? modalState(base) : null;
    let picture: ReturnType<typeof extractImage> = null;
    let signals: ReturnType<typeof parseSignals> = null;
    if (!modal) { // an open browser dialog blocks both screenshots and scripts; its text is already in the error
      const quiet = async (name: string, toolArgs: Record<string, unknown>) =>
        client.callTool({ name, arguments: toolArgs }, undefined, { timeout: 8_000 }).then(output => output as ToolOutput, () => null);
      const shot = await quiet("browser_take_screenshot", { type: "jpeg", filename: "live-view.jpeg" });
      picture = shot && !shot.isError ? extractImage(shot) : null;
      signals = parseSignals((await quiet("browser_evaluate", { function: SIGNALS_FUNCTION })) ?? undefined);
    }
    const text = buildDiagnostics({ tool, args, kind: info.kind, attempts: info.attempts, reason: info.reason, signals, modal, pictureAttached: Boolean(picture) });
    return {
      ...(base ?? {}), isError: true,
      content: [...(base?.content ?? []), ...(picture ? [{ type: "image", data: picture.data, mimeType: picture.mimeType }] : []), { type: "text", text }],
    };
  }

  private async browserCall(sessionId: string, tool: string, args: Record<string, unknown>, signal?: AbortSignal, onFrame?: FrameSink, turn?: string): Promise<unknown> {
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
      const guarded = GUARDED_TOOLS.has(tool);
      const refusal = guarded ? this.guard.check(sessionId, tool, args, turn) : null;
      if (refusal) {
        const refused = await this.explain(client, tool, args, null, { kind: "blocked", attempts: refusal.attempts, reason: refusal.reason });
        if (onFrame) this.scheduleFrame(sessionId, client, tool, args, refused, onFrame, "Blocked");
        return refused;
      }
      let result: ToolOutput;
      try {
        result = await client.callTool({ name: tool, arguments: args }, undefined, { timeout: 60_000, signal }) as ToolOutput;
      } catch (error) {
        if (guarded && !signal?.aborted) this.guard.record(sessionId, tool, args, false, turn); // a timeout counts as a failure too
        throw error;
      }
      const failed = guarded && result.isError === true;
      if (guarded) {
        const counts = this.guard.record(sessionId, tool, args, !failed, turn);
        if (failed) result = await this.explain(client, tool, args, result, { kind: "failed", attempts: counts.same });
      }
      if (onFrame) this.scheduleFrame(sessionId, client, tool, args, result, onFrame, failed ? "Failed" : undefined);
      return result;
    } finally { signal?.removeEventListener("abort", cancel); }
  }

  private scheduleFrame(sessionId: string, client: Client, tool: string, args: Record<string, unknown>, result: ToolOutput, onFrame: FrameSink, outcome?: "Failed" | "Blocked"): void {
    const state = this.frames.get(sessionId) ?? { seq: 0, token: 0, shown: false };
    this.frames.set(sessionId, state);
    const info = parsePageInfo(result);
    if (info.url) { state.url = info.url; state.title = info.title; }
    const label = outcome ? `${outcome}: ${describeBrowserAction(tool, args)}` : describeBrowserAction(tool, args);
    const emit = (image: { mimeType: "image/jpeg" | "image/png"; data: string }) => {
      if (image.data.length > BROWSER_FRAME_MAX_BASE64) return;
      state.shown = true;
      try { onFrame({ seq: state.seq++, label, url: state.url, title: state.title, ...image }); } catch { /* a closed pane is not an error for the assistant */ }
    };
    // When the assistant took a screenshot itself, that picture is the frame.
    const own = extractImage(result);
    if (own) { state.token++; emit(own); return; }
    if (result.isError || !(FRAME_AFTER.has(tool) || (FRAME_ON_FIRST_READ.has(tool) && !state.shown))) return;
    // Rapid actions are merged: only the newest one is photographed, after the earlier calls have finished.
    const token = ++state.token;
    const job = (this.queues.get(sessionId) ?? Promise.resolve()).then(async () => {
      if (state.token !== token || this.frames.get(sessionId) !== state) return;
      const shot = await client.callTool({ name: "browser_take_screenshot", arguments: { type: "jpeg", filename: "live-view.jpeg" } }, undefined, { timeout: 8_000 }) as ToolOutput;
      const image = shot.isError ? null : extractImage(shot);
      if (image && state.token === token) emit(image);
    }).catch(() => undefined);
    this.queues.set(sessionId, job);
  }

  async closeSession(sessionId: string, service?: Service): Promise<void> {
    if (!service) { this.profiles.delete(sessionId); this.queues.delete(sessionId); }
    if (!service || service === "browser") { this.frames.delete(sessionId); this.guard.clear(sessionId); }
    const pending = [...this.clients].filter(([key]) => service ? key === `${sessionId}:${service}` : key.startsWith(`${sessionId}:`))
      .map(([key, connection]) => ({ key, connection, outputDir: this.outputDirs.get(key) }));
    for (const { key } of pending) this.clients.delete(key);
    await Promise.allSettled(pending.map(async ({ connection }) => (await connection).close()));
    for (const { key, outputDir } of pending) this.cleanOutput(key, outputDir);
  }

  async close(): Promise<void> {
    this.profiles.clear();
    const sessions = new Set([...this.clients.keys()].map(key => key.slice(0, key.lastIndexOf(":"))));
    await Promise.allSettled([...sessions].map(sessionId => this.closeSession(sessionId)));
  }
}
