import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import spawn from "cross-spawn";
import { QaError } from "./util.js";

export interface Credentials { user: string; password: string }

/** Where encrypted passwords live: next to the bridge's own data, never inside a project (and so never in git). */
export function vaultDir(env: NodeJS.ProcessEnv = process.env): string {
  return env.QA_VAULT_DIR || path.join(env.LOCALAPPDATA || path.join(os.homedir(), ".config"), "Excel Partner", "qa-secrets");
}

export function vaultFile(origin: string, profile: string, env: NodeJS.ProcessEnv = process.env): string {
  const key = createHash("sha256").update(origin).digest("hex").slice(0, 16);
  return path.join(vaultDir(env), `${key}-${profile}.dpapi`);
}

export const envNames = (profile: string) => {
  const stem = `QA_${profile.toUpperCase().replace(/[^A-Z0-9]/g, "_")}`;
  return { user: `${stem}_USER`, password: `${stem}_PASSWORD` };
};

const POWERSHELL = process.env.SystemRoot ? path.join(process.env.SystemRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe") : "powershell.exe";

function powershell(script: string, env: Record<string, string>, stdin?: string): Promise<{ code: number | null; out: string; err: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(POWERSHELL, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
      { env: { ...process.env, ...env }, windowsHide: true, stdio: ["pipe", "pipe", "pipe"] });
    let out = "", err = "";
    child.stdout?.on("data", chunk => { out += chunk; });
    child.stderr?.on("data", chunk => { err += chunk; });
    child.once("error", reject);
    child.once("close", code => resolve({ code, out, err }));
    child.stdin?.on("error", () => undefined);
    child.stdin?.end(stdin ?? "");
  });
}

const PROTECT = "$in = [Console]::In.ReadToEnd(); ConvertTo-SecureString $in -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -LiteralPath $env:QA_VAULT_FILE -Encoding ascii";
const UNPROTECT = "$s = ConvertTo-SecureString ((Get-Content -LiteralPath $env:QA_VAULT_FILE -Raw).Trim()); [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))";

export const dpapiAvailable = (): boolean => process.platform === "win32";

/** Encrypt with Windows DPAPI (current user only). The secret reaches PowerShell on stdin, never on a command line. */
export async function storeCredentials(origin: string, profile: string, credentials: Credentials): Promise<string> {
  if (!dpapiAvailable()) throw new QaError(`Password storage uses Windows DPAPI. On this system set ${envNames(profile).user} and ${envNames(profile).password} instead.`, "unsupported");
  if (!credentials.user || !credentials.password) throw new QaError("A user name and a password are both required.", "invalid_arguments");
  const file = vaultFile(origin, profile);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const result = await powershell(PROTECT, { QA_VAULT_FILE: file }, JSON.stringify(credentials));
  if (result.code !== 0 || !fs.existsSync(file)) throw new QaError(`Could not encrypt the password: ${result.err.trim().split("\n")[0] || `exit ${result.code}`}`, "vault_error");
  try { fs.chmodSync(file, 0o600); } catch { /* best effort */ }
  return file;
}

export async function loadCredentials(origin: string, profile: string, env: NodeJS.ProcessEnv = process.env): Promise<Credentials | null> {
  const names = envNames(profile);
  if (env[names.user] && env[names.password]) return { user: env[names.user]!, password: env[names.password]! };
  const file = vaultFile(origin, profile);
  if (!dpapiAvailable() || !fs.existsSync(file)) return null;
  const result = await powershell(UNPROTECT, { QA_VAULT_FILE: file });
  if (result.code !== 0) throw new QaError(`Could not decrypt the stored password (it can only be read by the Windows user that saved it): ${result.err.trim().split("\n")[0]}`, "vault_error");
  try {
    const parsed = JSON.parse(result.out.trim()) as Partial<Credentials>;
    if (typeof parsed.user === "string" && typeof parsed.password === "string") return { user: parsed.user, password: parsed.password };
  } catch { /* fall through */ }
  throw new QaError("The stored password file is damaged. Run: excel-partner qa secret set <profile>", "vault_error");
}

export function hasCredentials(origin: string, profile: string, env: NodeJS.ProcessEnv = process.env): boolean {
  const names = envNames(profile);
  return Boolean((env[names.user] && env[names.password]) || fs.existsSync(vaultFile(origin, profile)));
}

export function clearCredentials(origin: string, profile: string): boolean {
  const file = vaultFile(origin, profile);
  if (!fs.existsSync(file)) return false;
  fs.rmSync(file, { force: true });
  return true;
}

/** Read one line from the terminal without echoing it. */
export function promptHidden(question: string, input: NodeJS.ReadStream = process.stdin, output: NodeJS.WriteStream = process.stderr): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!input.isTTY) { reject(new QaError("This command needs an interactive terminal.", "no_tty")); return; }
    output.write(question);
    let value = "";
    input.setRawMode(true); input.resume(); input.setEncoding("utf8");
    const done = (result: string | Error) => {
      input.setRawMode(false); input.pause(); input.removeListener("data", onData); output.write("\n");
      if (typeof result === "string") resolve(result); else reject(result);
    };
    const onData = (chunk: string) => {
      for (const char of chunk) {
        if (char === "\r" || char === "\n") return done(value);
        if (char === "\u0003") return done(new QaError("Cancelled.", "cancelled"));
        if (char === "\u007f" || char === "\b") value = value.slice(0, -1); else if (char >= " ") value += char;
      }
    };
    input.on("data", onData);
  });
}

export function promptLine(question: string, input: NodeJS.ReadStream = process.stdin, output: NodeJS.WriteStream = process.stderr): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!input.isTTY) { reject(new QaError("This command needs an interactive terminal.", "no_tty")); return; }
    output.write(question);
    input.resume(); input.setEncoding("utf8");
    const onData = (chunk: string) => { input.removeListener("data", onData); input.pause(); resolve(chunk.replace(/[\r\n]+$/, "")); };
    input.on("data", onData);
  });
}
