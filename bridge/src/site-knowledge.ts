import fs from "node:fs";
import path from "node:path";
import { normalizeOrigin } from "../../protocol/src/tools-testing.js";

export const KNOWLEDGE_KINDS = ["login", "navigation", "howto", "quirk", "data"] as const;
export type KnowledgeKind = (typeof KNOWLEDGE_KINDS)[number];
export const KNOWLEDGE_NAME_PATTERN = /^[a-z0-9][a-z0-9-]{0,47}$/;
export const MAX_ENTRIES_PER_SITE = 30;
export const MAX_CHARS_PER_SITE = 64_000;

export interface KnowledgeEntry { name: string; kind: KnowledgeKind; summary: string; updated: string; chars: number }

export class KnowledgeError extends Error {}

/** Things that must never be written down: the notes are read back into the model's context on every later visit. */
const SECRET_PATTERNS: Array<[RegExp, string]> = [
  [/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/, "a token"],
  [/\bBearer\s+[A-Za-z0-9._~+/=-]{16,}/i, "a bearer token"],
  [/\b(pass(word)?|pwd|passwd|mật\s*khẩu|secret|token|api[_-]?key|cookie|session[_-]?id)\s*[:=]\s*\S{3,}/i, "a password, token or cookie value"],
  [/-----BEGIN [A-Z ]*PRIVATE KEY-----/, "a private key"],
];

export function findSecret(text: string): string | null {
  for (const [pattern, what] of SECRET_PATTERNS) if (pattern.test(text)) return what;
  return null;
}

/** `https://staging.example.com:8443/web` -> `staging.example.com_8443`; null when it is not a web address. */
export function siteKey(origin: string): string | null {
  const normalized = normalizeOrigin(origin);
  if (!normalized) return null;
  const url = new URL(normalized);
  return `${url.hostname.replace(/[^a-z0-9.-]/gi, "-").toLowerCase()}${url.port ? `_${url.port}` : ""}`;
}

/** The reverse of siteKey, for screens that list sites by their folder name: `localhost_8081` -> `http://localhost:8081`. */
export function originOfKey(key: string): string | null {
  const match = /^([a-z0-9][a-z0-9.-]*?)(?:_(\d{1,5}))?$/.exec(key);
  if (!match || key.includes("..")) return null;
  const origin = `http://${match[1]}${match[2] ? `:${match[2]}` : ""}`;
  return siteKey(origin) === key ? origin : null;
}

const FRONT = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/;

function parse(text: string): { meta: Record<string, string>; body: string } {
  const match = FRONT.exec(text);
  if (!match) return { meta: {}, body: text };
  const meta: Record<string, string> = {};
  for (const line of match[1]!.split(/\r?\n/)) {
    const colon = line.indexOf(":");
    if (colon > 0) meta[line.slice(0, colon).trim()] = line.slice(colon + 1).trim();
  }
  return { meta, body: text.slice(match[0].length) };
}

/**
 * What the assistant has learned about each system, kept as plain Markdown files on this computer so the person can read,
 * edit and delete them. The index is not a separate file: it is read from each file's header, so a hand edit is never out of sync.
 */
export class SiteKnowledge {
  constructor(readonly dir: string) {}

  private siteDir(origin: string): string {
    const key = siteKey(origin);
    if (!key) throw new KnowledgeError(`'${origin}' is not a web address (use https://host, with the port if any).`);
    return path.join(this.dir, key);
  }

  private file(origin: string, name: string): string {
    if (!KNOWLEDGE_NAME_PATTERN.test(name)) throw new KnowledgeError(`'${name}' is not a valid note name (lowercase letters, digits and -, up to 48).`);
    return path.join(this.siteDir(origin), `${name}.md`);
  }

  list(origin: string): KnowledgeEntry[] {
    const dir = this.siteDir(origin);
    let files: string[] = [];
    try { files = fs.readdirSync(dir).filter(file => file.endsWith(".md")); } catch { return []; }
    const entries: KnowledgeEntry[] = [];
    for (const file of files) {
      const name = file.slice(0, -3);
      if (!KNOWLEDGE_NAME_PATTERN.test(name)) continue;
      try {
        const text = fs.readFileSync(path.join(dir, file), "utf8");
        const { meta, body } = parse(text);
        const kind = (KNOWLEDGE_KINDS as readonly string[]).includes(meta.kind ?? "") ? meta.kind as KnowledgeKind : "howto";
        entries.push({ name, kind, summary: (meta.summary ?? body.split(/\r?\n/).find(line => line.trim()) ?? "").slice(0, 200), updated: meta.updated ?? fs.statSync(path.join(dir, file)).mtime.toISOString().slice(0, 10), chars: body.length });
      } catch { /* unreadable file: left out of the index */ }
    }
    return entries.sort((a, b) => a.kind.localeCompare(b.kind) || a.name.localeCompare(b.name));
  }

  /** Every system that has notes, for `excel-partner knowledge` with no address. */
  sites(): string[] {
    try { return fs.readdirSync(this.dir, { withFileTypes: true }).filter(item => item.isDirectory()).map(item => item.name).sort(); } catch { return []; }
  }

  read(origin: string, name: string): { entry: KnowledgeEntry; content: string } {
    const file = this.file(origin, name);
    let text: string;
    try { text = fs.readFileSync(file, "utf8"); } catch { throw new KnowledgeError(`No note '${name}' for this site. Notes: ${this.list(origin).map(item => item.name).join(", ") || "none yet"}.` ); }
    const entry = this.list(origin).find(item => item.name === name);
    return { entry: entry ?? { name, kind: "howto", summary: "", updated: "", chars: text.length }, content: parse(text).body };
  }

  /** Check a note without writing it; save() repeats these checks. */
  validate(origin: string, input: { name: string; summary: string; content: string }): void {
    this.file(origin, input.name);
    const secret = findSecret(`${input.summary}\n${input.content}`);
    if (secret) throw new KnowledgeError(`The note looks like it contains ${secret}. Notes must never hold passwords, tokens or cookies: describe where to find them or what to do instead, and save again.`);
    const existing = this.list(origin);
    const others = existing.filter(item => item.name !== input.name);
    if (!existing.some(item => item.name === input.name) && existing.length >= MAX_ENTRIES_PER_SITE) {
      throw new KnowledgeError(`This site already has ${MAX_ENTRIES_PER_SITE} notes. Merge or delete some (delete_site_knowledge) before adding more.`);
    }
    if (others.reduce((sum, item) => sum + item.chars, 0) + input.content.length > MAX_CHARS_PER_SITE) {
      throw new KnowledgeError(`The notes for this site would exceed ${MAX_CHARS_PER_SITE} characters. Shorten or merge notes first.`);
    }
  }

  save(origin: string, input: { name: string; kind: KnowledgeKind; summary: string; content: string }, today = new Date()): KnowledgeEntry {
    this.validate(origin, input);
    const target = this.file(origin, input.name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    const updated = today.toISOString().slice(0, 10);
    const oneLine = (text: string) => text.replace(/\s+/g, " ").trim();
    const text = `---\nkind: ${input.kind}\nsummary: ${oneLine(input.summary)}\nupdated: ${updated}\n---\n${input.content.trim()}\n`;
    fs.writeFileSync(`${target}.tmp`, text, { mode: 0o600 });
    fs.renameSync(`${target}.tmp`, target);
    return { name: input.name, kind: input.kind, summary: oneLine(input.summary).slice(0, 200), updated, chars: input.content.trim().length };
  }

  delete(origin: string, name: string): boolean {
    try { fs.rmSync(this.file(origin, name)); return true; } catch (error) { if (error instanceof KnowledgeError) throw error; return false; }
  }
}
