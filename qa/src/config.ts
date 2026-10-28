import fs from "node:fs";
import path from "node:path";
import { z } from "zod";
import { ENV_RE, PROFILE_RE, QaError, readJson, writeFileAtomic } from "./util.js";

export const CONFIG_FILE = "qa.config.json";
export const AUTOMATION_DIR = "automation";

const Origin = z.string().url().refine(value => /^https?:\/\//.test(value) && new URL(value).origin === value.replace(/\/+$/, ""),
  "Use an origin only, such as https://staging.example.com (no path).").transform(value => value.replace(/\/+$/, ""));

/** A path on the site itself: "//host/x" or "/\host" would point somewhere else and carry the session cookie with it. */
const SitePath = z.string().regex(/^\/(?![/\\])/, "Use a path on the site, such as /web/login (not //host or a full URL).");

const LoginSchema = z.object({
  path: SitePath.default("/web/login").describe("Login page path (may include ?db=...)."),
  userSelector: z.string().default("input[name='login']"),
  passwordSelector: z.string().default("input[name='password']"),
  submitSelector: z.string().default("button[type='submit']"),
  /** Regex (as text) the URL must match once signed in. */
  successUrl: z.string().default("/(odoo|web)(?!/login)"),
}).strict();

const CheckSchema = z.object({
  type: z.enum(["odoo-session", "http-ok"]).default("odoo-session"),
  path: SitePath.default("/web/session/get_session_info"),
}).strict();

export const ProfileSchema = z.object({
  description: z.string().max(200).optional(),
  username: z.string().max(200).optional().describe("Not secret; the password is never stored here."),
  login: LoginSchema.default({}),
  check: CheckSchema.default({}),
}).strict();

export const DEFAULT_COLUMNS = {
  tcId: "TC_ID", reqId: "Req_ID", module: "Module", title: "Title", steps: "Steps", expected: "Expected",
  role: "Role", status: "Automation_Status", scriptPath: "Script_Path", lastRun: "Last_Run",
  lastResult: "Last_Result", duration: "Duration_s", evidence: "Run_Evidence",
} as const;

export const QaConfigSchema = z.object({
  version: z.literal(1).default(1),
  preset: z.enum(["odoo", "generic"]).default("odoo"),
  defaultEnvironment: z.string().regex(ENV_RE),
  environments: z.record(z.string().regex(ENV_RE), z.object({ baseUrl: Origin, ignoreHttpsErrors: z.boolean().default(false).describe("Accept a self-signed certificate (staging servers).") }).strict()).refine(value => Object.keys(value).length > 0, "Define at least one environment."),
  /** Hosts the browser may reach, besides each environment's own origin (for SSO providers, CDNs). */
  allowedOrigins: z.array(Origin).max(50).default([]),
  profiles: z.record(z.string().regex(PROFILE_RE), ProfileSchema).default({ default: {} }),
  workbook: z.object({
    testCasesSheet: z.string().min(1).max(31).default("TestCases"),
    columns: z.record(z.string().min(1).max(60)).default({}),
  }).strict().default({}),
  run: z.object({
    timeoutMs: z.number().int().min(5_000).max(600_000).default(60_000),
    workers: z.number().int().min(1).max(8).default(1),
    retries: z.number().int().min(0).max(3).default(1),
    totalTimeoutMinutes: z.number().int().min(1).max(240).default(30),
    keepDays: z.number().int().min(1).max(365).default(14),
    /** Let the assistant re-run a script it just repaired (one that had been approved) without asking each time, in Auto mode. */
    healWithoutPrompt: z.boolean().default(false),
  }).strict().default({}),
}).strict();

export type QaConfig = z.infer<typeof QaConfigSchema>;
export type QaProfile = z.infer<typeof ProfileSchema>;

export function columnsOf(config: QaConfig): Record<keyof typeof DEFAULT_COLUMNS, string> {
  return { ...DEFAULT_COLUMNS, ...config.workbook.columns } as Record<keyof typeof DEFAULT_COLUMNS, string>;
}

/** Every origin the browser may talk to: the environment origins plus the explicit extras. */
export function allowedOriginsOf(config: QaConfig): string[] {
  return [...new Set([...Object.values(config.environments).map(environment => environment.baseUrl), ...config.allowedOrigins])];
}

export function parseConfig(raw: unknown): QaConfig {
  const result = QaConfigSchema.safeParse(raw);
  if (!result.success) {
    const issue = result.error.issues[0]!;
    throw new QaError(`${CONFIG_FILE}: ${issue.path.join(".") || "(root)"}: ${issue.message}`, "invalid_config");
  }
  const config = result.data;
  if (!config.environments[config.defaultEnvironment]) throw new QaError(`${CONFIG_FILE}: defaultEnvironment '${config.defaultEnvironment}' is not defined in environments.`, "invalid_config");
  return config;
}

export function loadConfig(projectDir: string): QaConfig {
  const raw = readJson<unknown>(path.join(projectDir, CONFIG_FILE));
  if (!raw) throw new QaError(`No ${CONFIG_FILE} in ${projectDir}. Run: excel-partner qa init`, "no_project");
  return parseConfig(raw);
}

export function saveConfig(projectDir: string, config: QaConfig): void {
  writeFileAtomic(path.join(projectDir, CONFIG_FILE), JSON.stringify(config, null, 2) + "\n");
}

export function defaultConfig(options: { preset: "odoo" | "generic"; baseUrl: string; environment?: string }): QaConfig {
  const environment = options.environment ?? "staging";
  const generic = options.preset === "generic";
  return parseConfig({
    version: 1, preset: options.preset, defaultEnvironment: environment,
    environments: { [environment]: { baseUrl: options.baseUrl } },
    allowedOrigins: [],
    profiles: { admin: generic ? { login: { path: "/login", userSelector: "input[name='username']", passwordSelector: "input[name='password']", successUrl: "^(?!.*/login).*$" }, check: { type: "http-ok", path: "/" } } : {} },
  });
}

export function profileNames(config: QaConfig): string[] { return Object.keys(config.profiles); }

export function requireProfile(config: QaConfig, name: string): QaProfile {
  const profile = config.profiles[name];
  if (!profile) throw new QaError(`Unknown profile '${name}'. Defined: ${profileNames(config).join(", ") || "(none)"}. Add it to ${CONFIG_FILE}.`, "unknown_profile");
  return profile;
}

export function requireEnvironment(config: QaConfig, name?: string): { name: string; baseUrl: string } {
  const key = name ?? config.defaultEnvironment;
  const environment = config.environments[key];
  if (!environment) throw new QaError(`Unknown environment '${key}'. Defined: ${Object.keys(config.environments).join(", ")}.`, "unknown_environment");
  return { name: key, baseUrl: environment.baseUrl };
}

/** Locate the project that contains `start` (walks up to the nearest qa.config.json). */
export function findProject(start: string): string | null {
  let dir = path.resolve(start);
  for (let i = 0; i < 12; i++) {
    if (fs.existsSync(path.join(dir, CONFIG_FILE))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return null;
}
