import fs from "node:fs";
import os from "node:os";
import path from "node:path";

/**
 * Empty scratch directory used as the CLI working directory. Provider CLIs
 * are only meant to touch Excel through MCP, so they get no project files
 * (and no stray CLAUDE.md / AGENTS.md) to wander into.
 */
export function ensureWorkspaceDir(): string {
  const dir = path.join(process.env.EXCEL_AI_HOME ?? path.join(os.homedir(), ".cache", "excel-ai-addin"), "workspace");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}
