import fs from "node:fs";
import path from "node:path";
import spawn from "cross-spawn";
import { AUTOMATION_DIR, CONFIG_FILE, defaultConfig, saveConfig, type QaConfig } from "./config.js";
import { HARNESS_FILES, PLAYWRIGHT_VERSION, exampleTemplate, installHarness, installSkills, projectTemplate, templateWorkbookPath } from "./harness.js";
import { saveScript } from "./manifest.js";
import { QaError, writeFileAtomic } from "./util.js";

export interface InitOptions {
  dir: string;
  preset: "odoo" | "generic";
  baseUrl: string;
  environment?: string;
  /** Run npm install and download Chromium. */
  install?: boolean;
  /** Copy the /qa-* skills into the user's skill folder. */
  skills?: boolean;
  log?: (line: string) => void;
}

export function initProject(options: InitOptions): QaConfig {
  const dir = path.resolve(options.dir);
  if (fs.existsSync(path.join(dir, CONFIG_FILE))) throw new QaError(`${dir} is already a QA project (${CONFIG_FILE} exists).`, "exists");
  if (fs.existsSync(dir) && fs.readdirSync(dir).some(name => !name.startsWith("."))) throw new QaError(`${dir} is not empty. Choose a new or empty folder.`, "not_empty");
  const config = defaultConfig({ preset: options.preset, baseUrl: options.baseUrl, environment: options.environment });
  fs.mkdirSync(path.join(dir, AUTOMATION_DIR, "tests"), { recursive: true });
  fs.mkdirSync(path.join(dir, AUTOMATION_DIR, "pages"), { recursive: true });
  fs.mkdirSync(path.join(dir, "auth"), { recursive: true });
  saveConfig(dir, config);
  writeFileAtomic(path.join(dir, "package.json"), projectTemplate("package.json.tpl").replace("{{PLAYWRIGHT_VERSION}}", PLAYWRIGHT_VERSION));
  writeFileAtomic(path.join(dir, ".gitignore"), projectTemplate("gitignore"));
  writeFileAtomic(path.join(dir, "README.md"), projectTemplate("README.md"));
  installHarness(dir);
  fs.copyFileSync(templateWorkbookPath(), path.join(dir, "TestPlan.xlsx"));
  const role = Object.keys(config.profiles)[0]!;
  const example = exampleTemplate(options.preset === "odoo" ? "odoo.spec.js" : "generic.spec.js").replace("{{ROLE}}", role);
  const saved = saveScript(dir, config, { tcId: "TC_000", title: "Application loads", code: example, role, tags: ["smoke"] });
  if (!saved.ok) throw new QaError(`The example script failed validation: ${saved.violations[0]?.message}`, "internal");
  options.log?.(`Created ${dir}`);
  if (options.skills) {
    for (const item of installSkills()) options.log?.(`Skill /${item.name}: ${item.action}`);
  }
  if (options.install) installDependencies(dir, options.log);
  return config;
}

function run(command: string, args: string[], cwd: string, log?: (line: string) => void, quiet = false): void {
  log?.(`> ${command} ${args.join(" ")}`);
  const result = spawn.sync(command, args, { cwd, stdio: quiet ? "ignore" : "inherit", windowsHide: true });
  if (result.error) throw new QaError(`${command} could not start: ${result.error.message}`, "install_failed");
  if (result.status !== 0) throw new QaError(`${command} ${args[0]} failed (exit ${result.status}).`, "install_failed");
}

/** Install Playwright into the project (not into this package) and download the browser it needs. */
export function installDependencies(dir: string, log?: (line: string) => void, quiet = false): void {
  run("npm", ["install", "--no-audit", "--no-fund"], dir, log, quiet);
  run(process.execPath, [path.join(dir, "node_modules", "@playwright", "test", "cli.js"), "install", "chromium"], dir, log, quiet);
}

export const HARNESS_LIST = HARNESS_FILES;
