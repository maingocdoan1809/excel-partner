// Writes THIRD_PARTY_NOTICES.md: name, version, license and license text of every production dependency the package
// installs (npm-shrinkwrap.json, without dev packages). Runs in prepack, so the published package always carries it.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const lock = JSON.parse(fs.readFileSync(path.join(root, "npm-shrinkwrap.json"), "utf8"));
const entries = new Map();
for (const [location, info] of Object.entries(lock.packages)) {
  if (!location || info.dev || info.link) continue;
  const name = location.slice(location.lastIndexOf("node_modules/") + "node_modules/".length);
  const key = `${name}@${info.version}`;
  if (entries.has(key)) continue;
  const directory = path.join(root, location);
  const licenseFile = fs.existsSync(directory) ? fs.readdirSync(directory).find(file => /^(licen[cs]e|copying|notice)(\.|$)/i.test(file)) : undefined;
  entries.set(key, { name, version: info.version, license: info.license ?? "UNKNOWN", text: licenseFile ? fs.readFileSync(path.join(directory, licenseFile), "utf8").trim() : null });
}
const sorted = [...entries.values()].sort((a, b) => a.name.localeCompare(b.name) || a.version.localeCompare(b.version));
const out = ["# Third-party notices", "", "Office Partner includes or installs the following open-source packages. Provider CLIs (Claude Code, GitHub Copilot CLI, Codex, OpenCode) are not part of this package: they are downloaded from npm on the user's machine under their own licenses and terms.", ""];
for (const item of sorted) {
  out.push(`## ${item.name} ${item.version}`, "", `License: ${item.license}`, "");
  if (item.text) out.push("```", item.text, "```", "");
}
fs.writeFileSync(path.join(root, "THIRD_PARTY_NOTICES.md"), out.join("\n"));
console.log(`THIRD_PARTY_NOTICES.md: ${sorted.length} packages`);
