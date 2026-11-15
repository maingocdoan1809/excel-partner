import fs from "node:fs/promises";
import { createReadStream } from "node:fs";
import { createInterface } from "node:readline";
import os from "node:os";
import path from "node:path";
import type { StoredMessage, StoredTool } from "../../protocol/src/history.js";
import { ensureWorkspaceDir } from "./workspace.js";

/** How many conversations a provider lists, newest first. */
export const MAX_LISTED_SESSIONS = 200;
/** Same cut the pane applies when it saves a chat, so an imported one is not larger than a chat made here. */
const MAX_TEXT = 20_000;
const MAX_TOOL_TEXT = 4_000;
const MAX_TOOL_ARGS = 2_000;
const MAX_TITLE = 70;

export const userHome = (env: NodeJS.ProcessEnv): string => env.USERPROFILE ?? env.HOME ?? os.homedir();

/**
 * Where a conversation is read from and where the add-in's CLI looks for it. They differ when the product runs its own
 * copy of the CLI with a private state folder: a conversation from the person's terminal or VS Code lives in the CLI's
 * usual home, and an import has to put what the private copy needs to resume it into its own.
 */
export interface SessionHomes {
  source: NodeJS.ProcessEnv;
  target: NodeJS.ProcessEnv;
}

const clip = (text: string, max: number): string => (text.length > max ? `${text.slice(0, max)}…` : text);

export function samePath(a: string, b: string): boolean {
  const norm = (value: string) => path.resolve(value).replace(/[\\/]+$/, "");
  return process.platform === "win32" ? norm(a).toLowerCase() === norm(b).toLowerCase() : norm(a) === norm(b);
}

/** A conversation held in the scratch folder the add-in runs its CLIs in is the add-in's own: it is already in History. */
export const isOwnWorkspace = (cwd: string | undefined): boolean => Boolean(cwd) && samePath(cwd!, ensureWorkspaceDir());

export function titleFromText(text: string | undefined): string {
  const line = (text ?? "").replace(/\s+/g, " ").trim();
  return line.length > MAX_TITLE ? `${line.slice(0, MAX_TITLE - 3)}…` : line;
}

export function toMillis(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value < 1e12 ? value * 1000 : value;
  if (typeof value === "string") { const parsed = Date.parse(value); return Number.isNaN(parsed) ? undefined : parsed; }
  return undefined;
}

/**
 * Calls `onLine` with every parsed JSON line of a file, reading at most `maxBytes` of it from `start`. A line that is cut
 * off or damaged is skipped. Return false from `onLine` to stop reading.
 */
export async function scanJsonLines(file: string, onLine: (value: any) => boolean | void, range: { start?: number; maxBytes?: number } = {}): Promise<void> {
  const start = range.start ?? 0;
  const stream = createReadStream(file, { encoding: "utf8", start, ...(range.maxBytes ? { end: start + range.maxBytes - 1 } : {}) });
  const lines = createInterface({ input: stream, crlfDelay: Infinity });
  try {
    for await (const line of lines) {
      if (!line.trim()) continue;
      let value: unknown;
      try { value = JSON.parse(line); } catch { continue; }
      if (onLine(value) === false) break;
    }
  } finally { lines.close(); stream.destroy(); }
}

/** Newest `limit` entries (by modification time) from a list of files, ready to be summarised. */
export function newestFirst<T extends { mtime: number }>(items: T[], limit = MAX_LISTED_SESSIONS): T[] {
  return items.sort((a, b) => b.mtime - a.mtime).slice(0, limit);
}

export async function mtimeOf(file: string): Promise<number | undefined> {
  try { return (await fs.stat(file)).mtimeMs; } catch { return undefined; }
}

/** Builds the pane's messages from a provider's transcript: what was said and which tools ran, with results matched to calls. */
export class TranscriptBuilder {
  private readonly messages: StoredMessage[] = [];
  private readonly tools = new Map<string, StoredTool>();

  user(text: string | undefined): void {
    const value = text?.trim();
    if (value) this.messages.push({ kind: "user", text: clip(value, MAX_TEXT) });
  }

  /** Text of one reply: consecutive pieces (a CLI writes one per content block) read as one message. */
  assistant(text: string | undefined): void {
    const value = text?.trim();
    if (!value) return;
    const last = this.messages.at(-1);
    if (last?.kind === "assistant") last.text = clip(`${last.text}\n\n${value}`, MAX_TEXT);
    else this.messages.push({ kind: "assistant", text: clip(value, MAX_TEXT) });
  }

  toolCall(id: string, name: string, input?: unknown): void {
    let args = "";
    if (typeof input === "string") args = input;
    else if (input !== undefined) { try { args = JSON.stringify(input); } catch { /* A value that cannot be shown stays blank. */ } }
    const tool: StoredTool = { name, args: clip(args, MAX_TOOL_ARGS), state: "ok", error: null };
    const last = this.messages.at(-1);
    if (last?.kind === "tool_group") last.tools!.push(tool);
    else this.messages.push({ kind: "tool_group", text: "", tools: [tool] });
    this.tools.set(id, tool);
  }

  toolResult(id: string, output: string | undefined, failed = false): void {
    const tool = this.tools.get(id);
    if (!tool) return;
    if (output) tool.result = clip(output, MAX_TOOL_TEXT);
    if (failed) { tool.state = "error"; tool.error = clip(output ?? "The tool failed.", 500); }
  }

  get count(): number { return this.messages.filter((message) => message.kind === "user" || message.kind === "assistant").length; }

  build(): StoredMessage[] { return this.messages; }
}
