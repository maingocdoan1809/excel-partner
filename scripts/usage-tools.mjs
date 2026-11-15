// Which tools turns call most, what they return, and what the tool catalog costs per turn, from the local usage ledger.
// Usage: node scripts/usage-tools.mjs [path/to/usage.jsonl] [--sort calls|chars]
// Reads names and sizes only; the ledger holds no arguments or document contents.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const args = process.argv.slice(2);
const sortIndex = args.indexOf("--sort");
const sortBy = sortIndex >= 0 ? args.splice(sortIndex, 2)[1] : "calls";
const file = args[0] ?? path.join(process.env.OFFICE_PARTNER_DATA_DIR || path.join(process.env.LOCALAPPDATA || os.homedir(), "Excel Partner"), "usage.jsonl");

let text;
try { text = fs.readFileSync(file, "utf8"); } catch (error) { console.error(`Cannot read ${file}: ${error.code ?? error.message}`); process.exit(1); }

const tools = new Map();
let turns = 0, withTools = 0, catalogSum = 0, catalogTurns = 0;
for (const line of text.split("\n")) {
  if (!line.trim()) continue;
  let entry;
  try { entry = JSON.parse(line); } catch { continue; }
  if (entry.total || typeof entry.requestId !== "string") continue;
  turns++;
  if (typeof entry.catalogChars === "number") { catalogSum += entry.catalogChars; catalogTurns++; }
  if (!entry.tools) continue;
  withTools++;
  for (const [name, stats] of Object.entries(entry.tools)) {
    const total = tools.get(name) ?? { calls: 0, errors: 0, chars: 0, turns: 0 };
    total.calls += stats.calls; total.errors += stats.errors; total.chars += stats.resultChars; total.turns++;
    tools.set(name, total);
  }
}

console.log(`${file}\n${turns} turns, ${withTools} with tool calls.`);
if (catalogTurns) console.log(`Tool catalog offered per turn: ~${Math.round(catalogSum / catalogTurns)} chars (~${Math.round(catalogSum / catalogTurns / 4)} tokens) on average.`);
if (!tools.size) { console.log("No per-tool data yet: it is recorded for turns made by this version onward."); process.exit(0); }
const rows = [...tools].sort((a, b) => sortBy === "chars" ? b[1].chars - a[1].chars : b[1].calls - a[1].calls);
console.log(`\n${"tool".padEnd(30)} ${"calls".padStart(6)} ${"turns".padStart(6)} ${"errors".padStart(6)} ${"result chars".padStart(13)} ${"avg".padStart(8)}`);
for (const [name, t] of rows) console.log(`${name.padEnd(30)} ${String(t.calls).padStart(6)} ${String(t.turns).padStart(6)} ${String(t.errors).padStart(6)} ${String(t.chars).padStart(13)} ${String(Math.round(t.chars / t.calls)).padStart(8)}`);
