import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import { type QaConfig, type QaProfile, requireEnvironment, requireProfile } from "./config.js";
import { QaError, isInside } from "./util.js";

export interface SessionStatus { profile: string; environment: string; valid: boolean; reason?: string; savedAt?: string; unreachable?: boolean }

interface StoredCookie { name: string; value: string; domain: string; path: string; expires?: number; secure?: boolean }

export const authDir = (projectDir: string, environment: string): string => path.join(projectDir, "auth", environment);
export const authFile = (projectDir: string, environment: string, profile: string): string => path.join(authDir(projectDir, environment), `${profile}.json`);

export function readStoredCookies(file: string): StoredCookie[] | null {
  try {
    const state = JSON.parse(fs.readFileSync(file, "utf8")) as { cookies?: StoredCookie[] };
    return Array.isArray(state.cookies) ? state.cookies : [];
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw new QaError(`${path.basename(file)} is not a readable saved session. Run: excel-partner qa login`, "invalid_session");
  }
}

function cookieHeader(cookies: StoredCookie[], url: URL): string {
  const now = Date.now() / 1000;
  return cookies.filter(cookie => {
    if (cookie.expires && cookie.expires > 0 && cookie.expires < now) return false;
    const domain = cookie.domain.replace(/^\./, "").toLowerCase();
    const host = url.hostname.toLowerCase();
    if (!(host === domain || host.endsWith(`.${domain}`))) return false;
    if (cookie.secure && url.protocol !== "https:") return false;
    return url.pathname.startsWith(cookie.path || "/");
  }).map(cookie => `${cookie.name}=${cookie.value}`).join("; ");
}

interface Reply { status: number; location?: string; body: string }

function request(url: URL, options: { method: string; headers: Record<string, string>; body?: string; insecure: boolean; timeoutMs: number }): Promise<Reply> {
  return new Promise((resolve, reject) => {
    const client = url.protocol === "https:" ? https : http;
    const req = client.request(url, { method: options.method, headers: options.headers, timeout: options.timeoutMs, rejectUnauthorized: !options.insecure }, res => {
      const chunks: Buffer[] = [];
      let size = 0;
      res.on("data", (chunk: Buffer) => { size += chunk.length; if (size < 256_000) chunks.push(chunk); });
      res.on("end", () => resolve({ status: res.statusCode ?? 0, location: res.headers.location, body: Buffer.concat(chunks).toString("utf8") }));
    });
    req.on("timeout", () => req.destroy(new Error("timed out")));
    req.on("error", reject);
    req.end(options.body);
  });
}

/**
 * Asks the application whether a saved session still works. The cookie values are sent only to the
 * environment's own origin from the config, and nothing from the reply except the verdict is kept.
 */
export async function checkSession(projectDir: string, config: QaConfig, environmentName: string, profileName: string, timeoutMs = 10_000): Promise<SessionStatus> {
  const environment = requireEnvironment(config, environmentName);
  const profile: QaProfile = requireProfile(config, profileName);
  const base = { profile: profileName, environment: environment.name };
  const file = authFile(projectDir, environment.name, profileName);
  if (!isInside(projectDir, file)) throw new QaError("Invalid profile name.", "invalid_arguments");
  const cookies = readStoredCookies(file);
  if (!cookies) return { ...base, valid: false, reason: "no saved session" };
  const savedAt = fs.statSync(file).mtime.toISOString();
  const url = new URL(profile.check.path, environment.baseUrl);
  if (url.origin !== environment.baseUrl) throw new QaError("The session check path leaves the environment's origin.", "invalid_config");
  const cookie = cookieHeader(cookies, url);
  if (!cookie) return { ...base, valid: false, savedAt, reason: "saved session has no live cookie for this site" };
  const odoo = profile.check.type === "odoo-session";
  try {
    const reply = await request(url, {
      method: odoo ? "POST" : "GET",
      headers: { Cookie: cookie, Accept: "application/json, text/html", ...(odoo ? { "Content-Type": "application/json" } : {}) },
      body: odoo ? JSON.stringify({ jsonrpc: "2.0", method: "call", params: {} }) : undefined,
      insecure: config.environments[environment.name]!.ignoreHttpsErrors, timeoutMs,
    });
    if (odoo) {
      let uid: unknown;
      try { uid = (JSON.parse(reply.body) as { result?: { uid?: unknown } }).result?.uid; } catch { /* not JSON */ }
      return typeof uid === "number" && uid > 0 ? { ...base, valid: true, savedAt } : { ...base, valid: false, savedAt, reason: "the server no longer recognises this session" };
    }
    if (reply.status >= 300 && reply.status < 400) {
      return /login|signin|sso/i.test(reply.location ?? "") ? { ...base, valid: false, savedAt, reason: "redirected to the sign-in page" } : { ...base, valid: true, savedAt };
    }
    return reply.status >= 200 && reply.status < 300 ? { ...base, valid: true, savedAt } : { ...base, valid: false, savedAt, reason: `server answered ${reply.status}` };
  } catch (error) {
    return { ...base, valid: false, savedAt, unreachable: true, reason: `cannot reach ${url.host}: ${(error as Error).message}` };
  }
}
