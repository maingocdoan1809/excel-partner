import fs from "node:fs";
import path from "node:path";
import { AUTOMATION_DIR, type QaConfig, allowedOriginsOf } from "./config.js";
import { QaError, assertTcId, readJson, resolveInside, sha256Text, slugify, writeFileAtomic } from "./util.js";
import { compareAssertions, contentHash, validateFile, validateSource, type ValidationResult } from "./validator.js";

export const MANIFEST_FILE = "qa.manifest.json";

/** How many times the assistant may repair one script before a person has to step in (see saveScript). */
export const MAX_HEAL_ATTEMPTS = 2;

export type Health = "ok" | "failing" | "flaky" | "unknown";
/** What the runner may do with a script right now. `modified` means an approval exists but the content moved on. */
export type EffectiveStatus = "draft" | "approved" | "modified" | "invalid" | "missing";

export interface ScriptEntry {
  file: string;              // project-relative, posix
  title: string;
  req?: string;
  role: string;
  tags: string[];
  deps: string[];
  hash: string;              // combined hash of the current content
  savedAt: string;
  approved?: { by: string; at: string; hash: string };
  lastRun?: { runId: string; at: string; status: string; durationMs: number; health: Health };
  /** Repairs saved since the script last passed, was approved, or was reset by a person. */
  healAttempts?: number;
}

export interface Manifest {
  version: 1;
  harness?: { version: string; files: Record<string, string> };
  scripts: Record<string, ScriptEntry>;
}

export const automationDir = (projectDir: string): string => path.join(projectDir, AUTOMATION_DIR);
export const manifestPath = (projectDir: string): string => path.join(automationDir(projectDir), MANIFEST_FILE);

export function loadManifest(projectDir: string): Manifest {
  const manifest = readJson<Manifest>(manifestPath(projectDir));
  if (!manifest) return { version: 1, scripts: {} };
  if (manifest.version !== 1 || typeof manifest.scripts !== "object" || !manifest.scripts) throw new QaError(`${MANIFEST_FILE} has an unsupported format.`, "invalid_manifest");
  return manifest;
}

export function saveManifest(projectDir: string, manifest: Manifest): void {
  const sorted: Manifest = { ...manifest, scripts: Object.fromEntries(Object.entries(manifest.scripts).sort(([a], [b]) => a.localeCompare(b))) };
  writeFileAtomic(manifestPath(projectDir), JSON.stringify(sorted, null, 2) + "\n");
}

export interface Resolved { tcId: string; entry: ScriptEntry; status: EffectiveStatus; validation?: ValidationResult; currentHash?: string }

/** Re-validates and re-hashes a script from disk; the manifest alone is never trusted. */
export function resolveScript(projectDir: string, config: QaConfig, tcId: string, manifest = loadManifest(projectDir)): Resolved {
  const entry = manifest.scripts[tcId];
  if (!entry) throw new QaError(`No script is registered for ${tcId}.`, "unknown_script");
  const file = resolveInside(projectDir, entry.file);
  if (!fs.existsSync(file)) return { tcId, entry, status: "missing" };
  const validation = validateFile(file, { kind: "test", automationDir: automationDir(projectDir), allowedOrigins: allowedOriginsOf(config), expectedTcId: tcId });
  if (!validation.ok) return { tcId, entry, status: "invalid", validation };
  const currentHash = contentHash(projectDir, entry.file, validation.deps);
  const approved = entry.approved && entry.approved.hash === currentHash;
  return { tcId, entry, validation, currentHash, status: approved ? "approved" : entry.approved ? "modified" : "draft" };
}

export interface SaveInput { tcId: string; title: string; code: string; role?: string; tags?: string[] }
export interface SaveResult { ok: boolean; file?: string; status?: EffectiveStatus; violations: ValidationResult["violations"]; hash?: string }

/** Validate and store a test script as a draft. Any change voids a previous approval (the hash no longer matches). */
export function saveScript(projectDir: string, config: QaConfig, input: SaveInput, options: { dryRun?: boolean } = {}): SaveResult {
  assertTcId(input.tcId);
  const role = input.role ?? Object.keys(config.profiles)[0] ?? "default";
  if (!config.profiles[role]) throw new QaError(`Unknown role '${role}'. Defined profiles: ${Object.keys(config.profiles).join(", ")}.`, "unknown_profile");
  const manifest = loadManifest(projectDir);
  const previous = manifest.scripts[input.tcId];
  const relative = previous?.file ?? `${AUTOMATION_DIR}/tests/${input.tcId}_${slugify(input.title)}.spec.js`;
  const target = resolveInside(projectDir, relative);
  // Changing a script that had been approved, or that failed its last run, is a repair. Repairs are limited and may not weaken the checks.
  const repairing = Boolean(previous && (previous.approved || previous.lastRun?.health === "failing" || previous.lastRun?.health === "flaky"));
  if (repairing && (previous!.healAttempts ?? 0) >= MAX_HEAL_ATTEMPTS) {
    return { ok: false, violations: [{ rule: "repair-limit", message: `${input.tcId} has already been repaired ${previous!.healAttempts} times without passing. Stop and tell the person what you tried and what the page showed. They can allow more attempts with: excel-partner qa heal-reset ${input.tcId}` }] };
  }
  const validation = validateSource(input.code, { kind: "test", automationDir: automationDir(projectDir), file: target, allowedOrigins: allowedOriginsOf(config), expectedTcId: input.tcId });
  if (!validation.ok) return { ok: false, violations: validation.violations };
  if (repairing && fs.existsSync(target)) {
    const { removed } = compareAssertions(fs.readFileSync(target, "utf8"), input.code);
    if (removed.length) {
      return { ok: false, violations: removed.slice(0, 5).map(signature => ({ rule: "assertions", message: `A repair may not remove or change a check, but this one is gone: ${signature}. If the expected result really changed, a person must decide: update the test case and edit the script by hand. Otherwise keep the check and fix how the page is reached.` })) };
    }
  }
  if (options.dryRun) return { ok: true, file: relative, violations: [], status: previous?.approved ? "modified" : "draft" };
  writeFileAtomic(target, input.code.replace(/^﻿/, "").replace(/\r\n/g, "\n"));
  const hash = contentHash(projectDir, relative, validation.deps);
  manifest.scripts[input.tcId] = {
    file: relative, title: input.title.slice(0, 200), req: validation.reqId, role, tags: [...new Set((input.tags ?? []).map(tag => tag.replace(/^@/, "")))].slice(0, 10),
    deps: validation.deps, hash, savedAt: new Date().toISOString(),
    // An approval stays attached to its old hash so `status` can report "modified" instead of silently forgetting it.
    ...(previous?.approved ? { approved: previous.approved } : {}), ...(previous?.lastRun ? { lastRun: previous.lastRun } : {}),
    healAttempts: (previous?.healAttempts ?? 0) + (repairing ? 1 : 0),
  };
  saveManifest(projectDir, manifest);
  return { ok: true, file: relative, status: previous?.approved && previous.approved.hash === hash ? "approved" : previous?.approved ? "modified" : "draft", violations: [], hash };
}

/** Save a page object (shared helper). Voids approvals of every test that imports it, because their combined hash moves. */
export function savePage(projectDir: string, config: QaConfig, name: string, code: string, options: { dryRun?: boolean } = {}): SaveResult {
  if (!/^[a-z][a-z0-9-]{0,40}$/.test(name)) throw new QaError("Page name must be lower-case letters, digits or '-' (for example 'sale-order-page').", "invalid_arguments");
  const relative = `${AUTOMATION_DIR}/pages/${name}.js`;
  const target = resolveInside(projectDir, relative);
  const validation = validateSource(code, { kind: "page", automationDir: automationDir(projectDir), file: target, allowedOrigins: allowedOriginsOf(config) });
  if (!validation.ok) return { ok: false, violations: validation.violations };
  if (options.dryRun) return { ok: true, file: relative, violations: [] };
  writeFileAtomic(target, code.replace(/^﻿/, "").replace(/\r\n/g, "\n"));
  return { ok: true, file: relative, violations: [], hash: sha256Text(code) };
}

export function approve(projectDir: string, config: QaConfig, tcId: string, by: string): Resolved {
  const manifest = loadManifest(projectDir);
  const resolved = resolveScript(projectDir, config, tcId, manifest);
  if (resolved.status === "missing") throw new QaError(`${tcId}: the script file is missing.`, "missing_script");
  if (resolved.status === "invalid") throw new QaError(`${tcId}: fails validation (${resolved.validation!.violations[0]?.message}). Fix it before approving.`, "invalid_script");
  const entry = manifest.scripts[tcId]!;
  entry.deps = resolved.validation!.deps;
  entry.hash = resolved.currentHash!;
  entry.approved = { by, at: new Date().toISOString(), hash: resolved.currentHash! };
  entry.healAttempts = 0;
  // Keep what was approved, so the next approval can show exactly what changed since.
  writeFileAtomic(approvedSnapshotPath(projectDir, tcId), fs.readFileSync(resolveInside(projectDir, entry.file), "utf8").replace(/\r\n/g, "\n"));
  saveManifest(projectDir, manifest);
  return { ...resolved, entry, status: "approved" };
}

const approvedSnapshotPath = (projectDir: string, tcId: string): string => path.join(automationDir(projectDir), ".approved", `${assertTcId(tcId)}.spec.js.txt`);

/** The text of the script as it was when a person last approved it, or null if it was never approved. */
export function readApprovedSnapshot(projectDir: string, tcId: string): string | null {
  try { return fs.readFileSync(approvedSnapshotPath(projectDir, tcId), "utf8"); }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return null; throw error; }
}

/** A person allows the assistant more repair attempts for one script. */
export function resetHeal(projectDir: string, tcId: string): void {
  const manifest = loadManifest(projectDir);
  const entry = manifest.scripts[tcId];
  if (!entry) throw new QaError(`No script is registered for ${tcId}.`, "unknown_script");
  entry.healAttempts = 0;
  saveManifest(projectDir, manifest);
}

export function revoke(projectDir: string, tcId: string): void {
  const manifest = loadManifest(projectDir);
  const entry = manifest.scripts[tcId];
  if (!entry) throw new QaError(`No script is registered for ${tcId}.`, "unknown_script");
  delete entry.approved;
  saveManifest(projectDir, manifest);
}

export interface ListedScript { tcId: string; file: string; title: string; req?: string; role: string; tags: string[]; status: EffectiveStatus; health: Health; lastRun?: ScriptEntry["lastRun"]; healAttempts: number; problem?: string }

export function listScripts(projectDir: string, config: QaConfig): ListedScript[] {
  const manifest = loadManifest(projectDir);
  return Object.keys(manifest.scripts).sort().map(tcId => {
    const resolved = resolveScript(projectDir, config, tcId, manifest);
    const { entry } = resolved;
    return {
      tcId, file: entry.file, title: entry.title, req: entry.req, role: entry.role, tags: entry.tags, status: resolved.status,
      health: entry.lastRun?.health ?? "unknown", lastRun: entry.lastRun, healAttempts: entry.healAttempts ?? 0,
      problem: resolved.status === "invalid" ? resolved.validation!.violations[0]?.message : resolved.status === "missing" ? "script file is missing" : undefined,
    };
  });
}

export function recordRun(projectDir: string, runId: string, results: Array<{ tcId: string; status: string; durationMs: number; health: Health }>): void {
  const manifest = loadManifest(projectDir);
  const at = new Date().toISOString();
  for (const result of results) {
    const entry = manifest.scripts[result.tcId];
    if (!entry) continue;
    entry.lastRun = { runId, at, status: result.status, durationMs: result.durationMs, health: result.health };
    if (result.health === "ok") entry.healAttempts = 0; // it passes now, so whatever was repaired worked
  }
  saveManifest(projectDir, manifest);
}
