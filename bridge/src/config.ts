import path from "node:path";
import { defaultSkillDirs } from "./skills.js";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

/**
 * Repo root regardless of whether we run from `bridge/src/` (tsx) or
 * `dist/bridge/src/` (compiled): walk up until we find the package.json.
 */
function findProjectRoot(start: string): string {
  let dir = start;
  for (let i = 0; i < 10; i++) {
    if (fs.existsSync(path.join(dir, "package.json"))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return start;
}

export const PROJECT_ROOT = findProjectRoot(here);

function envInt(name: string, fallback: number): number {
  const raw = process.env[name];
  if (!raw) return fallback;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function envBool(name: string, fallback: boolean): boolean {
  const raw = process.env[name];
  if (raw === undefined) return fallback;
  return ["1", "true", "yes", "on"].includes(raw.toLowerCase());
}

/** Optional per-machine settings in `excel-partner.config.json` (not committed); environment variables win. */
export interface LocalConfig {
  /** Folders the assistant may read workbooks from (read-only). Windows or WSL paths. */
  referenceDirs?: string[];
  /** Extra folders searched for skills. */
  skillDirs?: string[];
  allowOfficeJs?: boolean;
  /** Extra hostnames the bridge answers to (e.g. a WSL IP); `excel-partner allow-host` edits this. */
  allowedHosts?: string[];
}

function readLocalConfig(): LocalConfig {
  try {
    const raw = JSON.parse(fs.readFileSync(process.env.BRIDGE_CONFIG ?? path.join(PROJECT_ROOT, "excel-partner.config.json"), "utf8")) as LocalConfig;
    const strings = (v: unknown) => (Array.isArray(v) ? v.filter((x): x is string => typeof x === "string" && x.trim() !== "") : []);
    return { referenceDirs: strings(raw.referenceDirs), skillDirs: strings(raw.skillDirs), allowedHosts: strings(raw.allowedHosts),
      allowOfficeJs: typeof raw.allowOfficeJs === "boolean" ? raw.allowOfficeJs : undefined };
  } catch {
    return {};
  }
}

const splitPaths = (value: string | undefined) => (value ?? "").split(path.delimiter).map((d) => d.trim()).filter(Boolean);

export interface BridgeConfig {
  host: string;
  port: number;
  /** Absolute paths to the TLS material used for the localhost HTTPS server. */
  tls: { certPath: string; keyPath: string } | null;
  /** Base URL the task pane and MCP processes should use to reach the bridge. */
  publicBaseUrl: string;
  /** Lower-case hostnames a request's Host header may carry; anything else (DNS rebinding) is refused. */
  allowedHosts: string[];
  sessionTtlMs: number;
  toolTimeoutMs: number;
  heartbeatMs: number;
  /** How long a provider gets to end a cancelled turn before its process is dropped. */
  cancelGraceMs: number;
  limits: {
    maxReadCells: number;
    maxWriteCells: number;
    maxSheets: number;
    maxSelectionCells: number;
  };
  /** Folders whose workbooks the assistant may read (read-only); empty = no file access. */
  referenceDirs: string[];
  /** Folders searched for skills (`<dir>/<skill>/SKILL.md`); earlier folders win on name clashes. */
  skillDirs: string[];
  /** Expose execute_office_js (arbitrary Office.js). On by default; can be disabled explicitly. */
  allowOfficeJs: boolean;
  logLevel: "debug" | "info" | "warn" | "error";
  /** Local settings file (config.json); the active QA project is recorded in it and re-read on every use. */
  configFile: string;
  /** Command used to launch the Excel MCP stdio server. */
  mcp: {
    command: string;
    args: string[];
  };
}

/** Loopback names, the host the bridge is configured for, and extras from the config file or BRIDGE_ALLOWED_HOSTS (e.g. a WSL IP). */
function defaultAllowedHosts(config: Pick<BridgeConfig, "host" | "publicBaseUrl">, extra: string[]): string[] {
  const names = ["localhost", "127.0.0.1", "::1", config.host];
  try { names.push(new URL(config.publicBaseUrl).hostname); } catch { /* a malformed public URL adds nothing */ }
  names.push(...extra, ...(process.env.BRIDGE_ALLOWED_HOSTS ?? "").split(",").map((name) => name.trim()));
  // 0.0.0.0 / :: are bind addresses, not names a browser sends in Host.
  return [...new Set(names.map((name) => name.toLowerCase().replace(/^\[(.*)\]$/, "$1")).filter((name) => name && name !== "0.0.0.0" && name !== "::"))];
}

export function loadConfig(overrides: Partial<BridgeConfig> = {}): BridgeConfig {
  const local = readLocalConfig();
  const host = process.env.BRIDGE_HOST ?? "127.0.0.1";
  const port = envInt("BRIDGE_PORT", 8765);
  const certPath = process.env.BRIDGE_CERT ?? path.join(PROJECT_ROOT, "certs", "localhost.crt");
  const keyPath = process.env.BRIDGE_KEY ?? path.join(PROJECT_ROOT, "certs", "localhost.key");
  const useTls = envBool("BRIDGE_TLS", true) && fs.existsSync(certPath) && fs.existsSync(keyPath);
  const scheme = useTls ? "https" : "http";

  const config: BridgeConfig = {
    host,
    port,
    tls: useTls ? { certPath, keyPath } : null,
    publicBaseUrl: process.env.BRIDGE_PUBLIC_URL ?? `${scheme}://${host}:${port}`,
    allowedHosts: [],
    sessionTtlMs: envInt("BRIDGE_SESSION_TTL_MS", 6 * 60 * 60 * 1000),
    toolTimeoutMs: envInt("BRIDGE_TOOL_TIMEOUT_MS", 120_000),
    heartbeatMs: envInt("BRIDGE_HEARTBEAT_MS", 20_000),
    cancelGraceMs: envInt("BRIDGE_CANCEL_GRACE_MS", 8_000),
    limits: {
      maxReadCells: envInt("BRIDGE_MAX_READ_CELLS", 20_000),
      maxWriteCells: envInt("BRIDGE_MAX_WRITE_CELLS", 50_000),
      maxSheets: envInt("BRIDGE_MAX_SHEETS", 500),
      maxSelectionCells: envInt("BRIDGE_MAX_SELECTION_CELLS", 2_000),
    },
    skillDirs: defaultSkillDirs(PROJECT_ROOT, [...(local.skillDirs ?? []), ...splitPaths(process.env.BRIDGE_SKILL_DIRS)]),
    referenceDirs: [...(local.referenceDirs ?? []), ...splitPaths(process.env.BRIDGE_REFERENCE_DIRS)],
    allowOfficeJs: envBool("BRIDGE_ALLOW_OFFICE_JS", local.allowOfficeJs ?? true),
    logLevel: (process.env.BRIDGE_LOG_LEVEL as BridgeConfig["logLevel"]) ?? "info",
    configFile: process.env.BRIDGE_CONFIG ?? path.join(PROJECT_ROOT, "excel-partner.config.json"),
    mcp: {
      command: process.env.BRIDGE_MCP_COMMAND ?? process.execPath,
      args: [
        process.env.BRIDGE_MCP_ENTRY ??
          path.join(PROJECT_ROOT, "dist", "excel_mcp", "src", "cli.js"),
      ],
    },
    ...overrides,
  };

  if (overrides.tls === undefined && overrides.publicBaseUrl === undefined && !useTls) {
    config.publicBaseUrl = `http://${host}:${port}`;
  }
  if (overrides.allowedHosts === undefined) config.allowedHosts = defaultAllowedHosts(config, local.allowedHosts ?? []);

  return config;
}
