import fs from "node:fs/promises";
import path from "node:path";
import type { ExternalSessionSummary, ImportedSession } from "../../../protocol/src/external-sessions.js";
import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { MAX_LISTED_SESSIONS, TranscriptBuilder, type SessionHomes, mtimeOf, newestFirst, scanJsonLines, titleFromText, toMillis, userHome } from "../external-sessions.js";
import { nativeTitle, validNativeId } from "../session-metadata.js";
import { ensureWorkspaceDir } from "../workspace.js";

/** Where Claude Code was run, by the `entrypoint` it records on every line. */
const ORIGINS: Record<string, string> = { cli: "Terminal", "claude-vscode": "VS Code", "claude-desktop": "Desktop app" };
const HEAD_BYTES = 1024 * 1024;
const TAIL_BYTES = 256 * 1024;

/** Claude Code keeps one folder per working directory, named after it; its conversations are the .jsonl files inside. */
export const claudeProjectsDir = (env: NodeJS.ProcessEnv): string => path.join(env.CLAUDE_CONFIG_DIR ?? path.join(userHome(env), ".claude"), "projects");
const projectFolder = (cwd: string): string => cwd.replace(/[^a-zA-Z0-9]/g, "-");

interface TranscriptFile { file: string; id: string; mtime: number }

async function transcriptFiles(env: NodeJS.ProcessEnv): Promise<TranscriptFile[]> {
  const root = claudeProjectsDir(env);
  const own = projectFolder(ensureWorkspaceDir());
  const found: TranscriptFile[] = [];
  let folders: string[] = [];
  try { folders = await fs.readdir(root); } catch { return found; }
  for (const folder of folders) {
    if (folder === own) continue;
    let names: string[] = [];
    try { names = await fs.readdir(path.join(root, folder)); } catch { continue; }
    for (const name of names) {
      const id = name.endsWith(".jsonl") ? name.slice(0, -6) : "";
      if (!id || !validNativeId(id)) continue;
      const file = path.join(root, folder, name);
      const mtime = await mtimeOf(file);
      if (mtime !== undefined) found.push({ file, id, mtime });
    }
  }
  return found;
}

/** The words a person typed on a line, without what Claude Code or the add-in wrapped around them. */
function humanText(line: any): string {
  if (line.type !== "user" || line.isSidechain || line.isMeta || line.isCompactSummary) return "";
  const content = line.message?.content;
  const text = typeof content === "string" ? content
    : Array.isArray(content) ? content.filter((block: any) => block?.type === "text" && typeof block.text === "string").map((block: any) => block.text).join("\n") : "";
  const cleaned = text.replace(/<system-reminder>[\s\S]*?<\/system-reminder>/g, "").trim();
  if (/^<(?:command-name|command-message|local-command-\w+|user-prompt-submit-hook)\b/.test(cleaned) || /^\[Request interrupted/.test(cleaned) || /^Caveat: /.test(cleaned)) return "";
  return cleaned;
}

async function summarize({ file, id, mtime }: TranscriptFile): Promise<ExternalSessionSummary | null> {
  let cwd: string | undefined, entrypoint: string | undefined, prompt = "", title: string | undefined;
  const take = (line: any) => {
    if (line.type === "custom-title") title = nativeTitle(line.customTitle) ?? title;
    else if (line.type === "ai-title") title = title ?? nativeTitle(line.aiTitle);
    else if (line.type === "summary") title = title ?? nativeTitle(line.summary);
    if (!prompt && line.type === "user") {
      const text = humanText(line);
      if (text) { prompt = text; cwd = typeof line.cwd === "string" ? line.cwd : cwd; entrypoint = typeof line.entrypoint === "string" ? line.entrypoint : entrypoint; }
    }
  };
  const size = (await fs.stat(file)).size;
  await scanJsonLines(file, take, { maxBytes: HEAD_BYTES });
  // A title the CLI wrote late in a long conversation sits near the end of the file.
  if (size > HEAD_BYTES) await scanJsonLines(file, take, { start: Math.max(HEAD_BYTES, size - TAIL_BYTES) });
  if (!prompt) return null;
  return { id, title: cleanSessionTitle(title) || titleFromText(prompt), origin: ORIGINS[entrypoint ?? ""] ?? "Claude Code", ...(cwd ? { cwd } : {}), updatedAt: mtime };
}

export async function listClaudeSessions(env: NodeJS.ProcessEnv): Promise<ExternalSessionSummary[]> {
  const files = newestFirst(await transcriptFiles(env), MAX_LISTED_SESSIONS);
  const summaries = await Promise.all(files.map((item) => summarize(item).catch(() => null)));
  return summaries.filter((item): item is ExternalSessionSummary => item !== null);
}

/**
 * Reads a Claude Code conversation and puts its transcript where the add-in's Claude looks: `--resume` finds a
 * conversation only in the folder of the directory it runs in, which is the add-in's scratch folder, not the one the
 * conversation was held in. The original is copied, never moved, so the terminal or VS Code can still open it.
 */
export async function importClaudeSession(homes: SessionHomes, id: string): Promise<ImportedSession> {
  if (!validNativeId(id)) throw new Error("Invalid Claude session ID");
  const source = (await transcriptFiles(homes.source)).find((item) => item.id === id);
  if (!source) throw new Error("That Claude conversation was not found on this computer.");
  const transcript = new TranscriptBuilder();
  let title: string | undefined, startedAt: number | undefined, updatedAt: number | undefined;
  await scanJsonLines(source.file, (line) => {
    const at = toMillis(line.timestamp);
    if (at !== undefined && line.type !== "file-history-snapshot") { startedAt ??= at; updatedAt = at; }
    if (line.type === "custom-title") title = nativeTitle(line.customTitle) ?? title;
    else if (line.type === "ai-title") title = title ?? nativeTitle(line.aiTitle);
    if (line.isSidechain || (line.type !== "user" && line.type !== "assistant")) return;
    const content = line.message?.content;
    if (line.type === "user") {
      transcript.user(humanText(line));
      if (Array.isArray(content)) for (const block of content) if (block?.type === "tool_result") {
        const output = typeof block.content === "string" ? block.content
          : Array.isArray(block.content) ? block.content.map((part: any) => typeof part?.text === "string" ? part.text : "").join("\n") : undefined;
        transcript.toolResult(String(block.tool_use_id), output, block.is_error === true);
      }
      return;
    }
    if (typeof content === "string") transcript.assistant(content);
    else if (Array.isArray(content)) for (const block of content) {
      if (block?.type === "text") transcript.assistant(block.text);
      else if (block?.type === "tool_use") transcript.toolCall(String(block.id), String(block.name ?? "tool"), block.input);
    }
  });
  const messages = transcript.build();
  const target = path.join(claudeProjectsDir(homes.target), projectFolder(ensureWorkspaceDir()));
  await fs.mkdir(target, { recursive: true });
  try { await fs.copyFile(source.file, path.join(target, `${id}.jsonl`), fs.constants.COPYFILE_EXCL); }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error; }
  const first = messages.find((message) => message.kind === "user")?.text;
  return { providerSessionId: id, title: cleanSessionTitle(title) || titleFromText(first), startedAt: startedAt ?? source.mtime, updatedAt: updatedAt ?? source.mtime, messages };
}
