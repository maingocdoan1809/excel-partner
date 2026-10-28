import { createHash, randomBytes } from "node:crypto";
import fs from "node:fs";
import path from "node:path";

export class QaError extends Error {
  constructor(message: string, readonly code: string = "qa_error") { super(message); }
}

export const sha256 = (data: string | Buffer): string => createHash("sha256").update(data).digest("hex");

/** Hash text with line endings normalised so a Windows checkout does not void approvals. */
export const sha256Text = (text: string): string => sha256(text.replace(/\r\n/g, "\n"));

export function isInside(root: string, target: string): boolean {
  const relative = path.relative(path.resolve(root), path.resolve(target));
  return relative !== "" && !relative.startsWith("..") && !path.isAbsolute(relative);
}

/** Resolve `relative` under `root`, refusing traversal, absolute paths and links that leave the root. */
export function resolveInside(root: string, relative: string): string {
  if (!relative || path.isAbsolute(relative) || /[\0:]/.test(relative) || relative.split(/[\\/]/).includes("..")) {
    throw new QaError(`'${relative}' is not a relative path inside the QA project.`, "invalid_path");
  }
  const target = path.resolve(root, relative);
  if (!isInside(root, target)) throw new QaError(`'${relative}' is outside the QA project.`, "invalid_path");
  // A link inside the project must not point outside it.
  let probe = target;
  while (!fs.existsSync(probe) && probe !== root) probe = path.dirname(probe);
  if (fs.existsSync(probe) && !(probe === path.resolve(root) || isInside(fs.realpathSync(root), fs.realpathSync(probe)))) {
    throw new QaError(`'${relative}' resolves outside the QA project.`, "invalid_path");
  }
  return target;
}

export function writeFileAtomic(file: string, content: string | Buffer, mode?: number): void {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const temporary = `${file}.${process.pid}.${randomBytes(4).toString("hex")}.tmp`;
  fs.writeFileSync(temporary, content, mode === undefined ? undefined : { mode });
  fs.renameSync(temporary, file);
}

export function readJson<T>(file: string): T | null {
  try { return JSON.parse(fs.readFileSync(file, "utf8")) as T; }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return null; throw error; }
}

export const slugify = (text: string, max = 40): string =>
  text.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/gi, "d")
    .toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, max).replace(/-+$/, "") || "test";

/** `TC_001`, `TC-SALES-12`, `LOGIN_01`: letters/digits in 2-4 dash- or underscore-separated groups. */
export const TC_ID_RE = /^[A-Za-z][A-Za-z0-9]{0,11}([_-][A-Za-z0-9]{1,12}){1,3}$/;
export function assertTcId(id: string): string {
  if (!TC_ID_RE.test(id)) throw new QaError(`'${id}' is not a valid test case id (use letters, digits, '_' or '-', e.g. TC_001).`, "invalid_tc_id");
  return id;
}

export const PROFILE_RE = /^[a-z0-9][a-z0-9_-]{0,31}$/;
export const ENV_RE = /^[a-z0-9][a-z0-9_-]{0,31}$/;

/** Variables a test process may inherit: enough to run, nothing that carries credentials. */
const SAFE_ENV = /^(path|pathext|systemroot|windir|comspec|home|userprofile|homedrive|homepath|temp|tmp|tmpdir|localappdata|appdata|programdata|programfiles(\(x86\))?|programw6432|commonprogramfiles(\(x86\))?|processor_architecture|number_of_processors|os|lang|lc_.*|tz|ms_playwright_browsers_path|playwright_browsers_path|node_options)$/i;
export function scrubbedEnv(extra: Record<string, string | undefined> = {}, source: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = {};
  for (const [key, value] of Object.entries(source)) if (value !== undefined && SAFE_ENV.test(key)) env[key] = value;
  // NODE_OPTIONS can preload arbitrary code; only keep it when it holds nothing but memory flags.
  if (env.NODE_OPTIONS && !/^(\s*--max-old-space-size=\d+\s*)+$/.test(env.NODE_OPTIONS)) delete env.NODE_OPTIONS;
  for (const [key, value] of Object.entries(extra)) if (value !== undefined) env[key] = value;
  return env;
}

/** Remove things that look like credentials from text that will be stored or shown. */
export function redact(text: string): string {
  return text
    .replace(/(authorization|proxy-authorization)(["']?\s*[:=]\s*["']?)(bearer|basic|token)?\s*[^\s"',;]+/gi, "$1$2[redacted]")
    .replace(/((?:set-)?cookie["']?\s*[:=]\s*["']?)[^\n"']+/gi, "$1[redacted]")
    .replace(/\b(session_id|sessionid|access_token|refresh_token|id_token|csrf_token|api[_-]?key|secret|password|passwd|pwd)(["']?\s*[:=]\s*["']?)[^\s"',;&]+/gi, "$1$2[redacted]")
    .replace(/\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/g, "[redacted-jwt]")
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]{12,}/gi, "Bearer [redacted]");
}
