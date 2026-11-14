import fs from "node:fs";
import path from "node:path";
import { LOCKED_DOWN_POLICY, OPEN_POLICY, parseAdminPolicy, type AdminPolicy, type ToolCapabilities } from "../../protocol/src/index.js";

/** The admin policy in force, where it was read from, and why it was locked down when the file could not be used. */
export interface LoadedPolicy { policy: AdminPolicy; file: string; error?: string }

export function policyFile(): string {
  return process.env.OFFICE_PARTNER_POLICY_FILE
    ?? path.join(process.env.ProgramData ?? process.env.PROGRAMDATA ?? "C:\\ProgramData", "Office Partner", "policy.json");
}

/** No file: no restrictions. A file that cannot be read or parsed: everything risky off (see LOCKED_DOWN_POLICY). */
export function loadAdminPolicy(file = policyFile()): LoadedPolicy {
  let text: string;
  try { text = fs.readFileSync(file, "utf8"); }
  catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === "ENOENT" || code === "ENOTDIR") return { policy: OPEN_POLICY, file };
    return { policy: LOCKED_DOWN_POLICY, file, error: `policy.json could not be read (${code ?? "unknown error"})` };
  }
  return { file, ...parseAdminPolicy(text.replace(/^\uFEFF/, "")) }; // Windows editors and PowerShell often write a BOM
}

/** The tools a pane asked for, minus what the policy forbids. */
export function allowedCapabilities(policy: AdminPolicy, requested: ToolCapabilities): ToolCapabilities {
  return { ...requested, webSearch: requested.webSearch && policy.allowWebSearch, playwright: requested.playwright && policy.allowBrowserAutomation };
}
