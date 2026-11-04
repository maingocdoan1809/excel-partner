import fs from "node:fs";
import path from "node:path";

export const SESSION_NAME_PATTERN = /^[a-z0-9][a-z0-9_-]{0,31}$/;

export interface TestSessionInfo { name: string; origins: string[]; savedAt: string; url: string }

/**
 * Signed-in browser sessions saved by `excel-partner login`. The state file holds cookies, which are credentials:
 * it stays on this computer, is copied into a temporary folder when a browser starts, and is never shown to the assistant.
 */
export class TestSessions {
  constructor(readonly dir: string) {}

  private file(name: string, kind: "state" | "meta"): string {
    if (!SESSION_NAME_PATTERN.test(name)) throw new Error(`'${name}' is not a valid session name (lowercase letters, digits, - and _; up to 32).`);
    return path.join(this.dir, `${name}.${kind}.json`);
  }

  get(name: string): { statePath: string; info: TestSessionInfo } | null {
    try {
      const info = JSON.parse(fs.readFileSync(this.file(name, "meta"), "utf8")) as TestSessionInfo;
      const statePath = this.file(name, "state");
      return fs.existsSync(statePath) && Array.isArray(info.origins) ? { statePath, info: { ...info, name } } : null;
    } catch { return null; }
  }

  list(): TestSessionInfo[] {
    let files: string[] = [];
    try { files = fs.readdirSync(this.dir); } catch { return []; }
    return files.filter(file => file.endsWith(".meta.json"))
      .map(file => this.get(file.slice(0, -".meta.json".length))?.info)
      .filter((info): info is TestSessionInfo => Boolean(info))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  save(name: string, state: unknown, info: Omit<TestSessionInfo, "name">): void {
    fs.mkdirSync(this.dir, { recursive: true });
    for (const [kind, body] of [["state", state], ["meta", { name, ...info }]] as const) {
      const target = this.file(name, kind);
      fs.writeFileSync(`${target}.tmp`, JSON.stringify(body, null, 2), { mode: 0o600 });
      fs.renameSync(`${target}.tmp`, target);
    }
  }

  remove(name: string): boolean {
    let removed = false;
    for (const kind of ["state", "meta"] as const) {
      try { fs.rmSync(this.file(name, kind)); removed = true; } catch { /* already gone */ }
    }
    return removed;
  }
}
