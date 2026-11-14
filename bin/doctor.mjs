#!/usr/bin/env node
// This entrypoint also works without a compiled bridge or third-party packages.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { collectDoctor, PROVIDERS } from "./doctor-checks.mjs";

const USAGE = "Usage: office-partner-doctor [addin|gateway|all] [--provider <id>] [--probe] [--mcp] [--integrity] [--host excel|word|powerpoint] [--json] [--output <new file>]\n";

export function parseDoctorArgs(args) {
  const result = { target: "all", probe: false, mcp: false, integrity: false, host: "excel", json: false, help: false };
  let targetSet = false;
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (["addin", "gateway", "all"].includes(arg) && !targetSet) { result.target = arg; targetSet = true; }
    else if (arg === "--help" || arg === "-h") result.help = true;
    else if (arg === "--json") result.json = true;
    else if (arg === "--probe") result.probe = true;
    else if (arg === "--mcp") result.mcp = true;
    else if (arg === "--integrity") result.integrity = true;
    else if (["--provider", "--host", "--output"].includes(arg)) {
      const value = args[++i];
      if (!value || value.startsWith("--")) throw new Error(`Missing value for ${arg}.`);
      result[arg.slice(2)] = value;
    } else throw new Error("Unknown doctor argument. Use doctor [addin|gateway|all] [--provider <id>] [--probe] [--mcp] [--host <host>] [--json] [--output <file>].");
  }
  if (result.provider && !PROVIDERS.includes(result.provider)) throw new Error("Provider must be claude, copilot, codex or opencode.");
  if (!["excel", "word", "powerpoint"].includes(result.host)) throw new Error("Host must be excel, word or powerpoint.");
  return result;
}

export async function doctorMain(args, options = {}) {
  const parsed = parseDoctorArgs(args);
  if (parsed.help) { (options.write ?? (text => process.stdout.write(text)))(USAGE); return 0; }
  const root = options.root ?? path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const report = await collectDoctor({ ...options, ...parsed, root });
  const serialized = JSON.stringify(report, null, 2) + "\n";
  if (parsed.output) {
    // Never overwrite a config, document or earlier report through this option.
    try { fs.writeFileSync(parsed.output, serialized, { flag: "wx", mode: 0o600 }); }
    catch { throw new Error("Could not save diagnostics. Choose a new file in an existing writable directory."); }
  }
  const write = options.write ?? (text => process.stdout.write(text));
  if (parsed.json) write(serialized);
  else {
    write(`Office Partner diagnostics (${report.product.name} ${report.product.version ?? "unknown"})\nReport: ${report.diagnosticId}\n`);
    for (const check of report.checks) {
      write(`[${check.status}] ${check.id}: ${check.code} — ${check.message}\n`);
      if (check.action && check.status !== "pass") write(`  ${check.action}\n`);
    }
    write(`Result: ${report.status}. Exit codes: 0 = no failures, 1 = failed checks, 2 = usage/internal error.\n`);
    if (parsed.output) write("Saved a JSON report containing diagnostic metadata only.\n");
  }
  return report.status === "fail" ? 1 : 0;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  doctorMain(process.argv.slice(2)).then(code => { process.exitCode = code; }, () => {
    // No exception dumps: even filesystem errors can contain private paths.
    process.stderr.write("Doctor could not run. Check the arguments and the installation of its three bin/doctor files.\n");
    process.exitCode = 2;
  });
}
