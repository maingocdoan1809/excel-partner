import fs from "node:fs";
import path from "node:path";
import { CONFIG_FILE } from "./config.js";
import { QaError, readJson, writeFileAtomic } from "./util.js";

/** The bridge may read and write only inside the project recorded here (`qaProject` in config.json). */
export function readActiveProject(configFile: string): string | null {
  const config = readJson<{ qaProject?: unknown }>(configFile);
  return typeof config?.qaProject === "string" && config.qaProject ? config.qaProject : null;
}

export function setActiveProject(configFile: string, projectDir: string | null): void {
  const absolute = projectDir ? path.resolve(projectDir) : null;
  if (absolute && !fs.existsSync(path.join(absolute, CONFIG_FILE))) throw new QaError(`${absolute} has no ${CONFIG_FILE}. Run: excel-partner qa init`, "no_project");
  const current = readJson<Record<string, unknown>>(configFile) ?? {};
  if (absolute) current.qaProject = absolute; else delete current.qaProject;
  writeFileAtomic(configFile, JSON.stringify(current, null, 2) + "\n", 0o600);
}
