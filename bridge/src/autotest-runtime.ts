import { execFile, spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { pythonInterpreter } from "./python-runner.js";

export const MIN_PYTHON = [3, 9] as const;

export interface RuntimeStatus {
  python: string | null;
  /** The kit needs Python 3.9 or newer. */
  pythonOk: boolean;
  missing: string[];
  ready: boolean;
  install: { state: "idle" | "running" | "done" | "failed"; log: string };
}

const MODULES: Array<[string, string]> = [["playwright", "playwright"], ["openpyxl", "openpyxl"], ["PIL", "pillow"]];
const LOG_TAIL = 2000;

/** The Python side of the kit: is it there, and install its libraries with pip (one run at a time). */
export class AutotestRuntime {
  private state: RuntimeStatus["install"]["state"] = "idle";

  constructor(
    private readonly home: () => string,
    private readonly python = pythonInterpreter(),
  ) {}

  private run(args: string[], timeout = 15_000): Promise<{ ok: boolean; out: string }> {
    return new Promise(resolve => {
      execFile(this.python, args, { timeout, windowsHide: true }, (error, stdout, stderr) => resolve({ ok: !error, out: `${stdout}${stderr}`.trim() }));
    });
  }

  private get logFile(): string { return path.join(this.home(), "logs", "runtime-install.log"); }

  private tail(): string {
    try { return fs.readFileSync(this.logFile, "utf8").slice(-LOG_TAIL); } catch { return ""; }
  }

  async check(): Promise<RuntimeStatus> {
    const version = await this.run(["--version"]);
    const python = version.ok ? version.out.replace(/^Python\s+/i, "") : null;
    const missing: string[] = [];
    if (python) {
      for (const [module, pip] of MODULES) if (!(await this.run(["-c", `import ${module}`])).ok) missing.push(pip);
    }
    const [major = 0, minor = 0] = (python ?? "").split(".").map(Number);
    const pythonOk = Boolean(python) && (major > MIN_PYTHON[0] || (major === MIN_PYTHON[0] && minor >= MIN_PYTHON[1]));
    return { python, pythonOk, missing: python ? missing : MODULES.map(([, pip]) => pip), ready: pythonOk && !missing.length, install: { state: this.state, log: this.tail() } };
  }

  /**
   * Build the testcase catalog from the testcases folder, so the assistant never starts from "no catalog". Best effort and silent:
   * a missing Python or library only means the assistant builds it itself. Output goes to logs/catalog.log.
   */
  buildCatalog(): void {
    const home = this.home();
    if (!fs.existsSync(path.join(home, "qakit", "__init__.py"))) return;
    try {
      fs.mkdirSync(path.join(home, "logs"), { recursive: true });
      const log = fs.openSync(path.join(home, "logs", "catalog.log"), "w");
      const child = spawn(this.python, ["-m", "qakit.tools.catalog", "build", "--note", "synced from the open workbook"], { cwd: home, windowsHide: true, stdio: ["ignore", log, log], env: { ...process.env, QA_HOME: home } });
      child.on("error", () => undefined);
      child.on("close", () => { try { fs.closeSync(log); } catch { /* already closed */ } });
    } catch { /* the assistant builds it itself */ }
  }

  /** pip install of the kit's requirements, then the browser Playwright drives. Returns at once; progress is in `check().install`. */
  install(): void {
    if (this.state === "running") return;
    const requirements = path.join(this.home(), "requirements.txt");
    if (!fs.existsSync(requirements)) throw new Error("Save the settings first: that installs the kit and its requirements list.");
    fs.mkdirSync(path.dirname(this.logFile), { recursive: true });
    fs.writeFileSync(this.logFile, "", "utf8");
    this.state = "running";
    const steps: string[][] = [["-m", "pip", "install", "-r", requirements], ["-m", "playwright", "install", "chromium"]];
    const next = (index: number) => {
      if (index >= steps.length) { this.state = "done"; return; }
      fs.appendFileSync(this.logFile, `\n$ python ${steps[index]!.join(" ")}\n`);
      const child = spawn(this.python, steps[index]!, { windowsHide: true, cwd: this.home() });
      const append = (chunk: Buffer) => { try { fs.appendFileSync(this.logFile, chunk); } catch { /* log only */ } };
      child.stdout.on("data", append); child.stderr.on("data", append);
      child.on("error", error => { append(Buffer.from(`\n${error.message}\n`)); this.state = "failed"; });
      child.on("close", code => { if (code === 0) next(index + 1); else if (this.state !== "failed") { append(Buffer.from(`\nexit ${code}\n`)); this.state = "failed"; } });
    };
    next(0);
  }
}
