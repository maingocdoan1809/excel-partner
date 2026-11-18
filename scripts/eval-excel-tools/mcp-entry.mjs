// MCP entry for variant runs: applies the variant's tool-description change, then starts the normal Excel MCP server.
// `v1` is the audit's proposed insert_rows description (see .claude/hillclimb/excel-tools/v1/change.md).
import path from "node:path";
import { pathToFileURL } from "node:url";

const root = path.resolve(import.meta.dirname, "..", "..");
const defs = await import(pathToFileURL(path.join(root, "dist", "shared", "src", "tool-definitions.js")).href);
const V1 = "WRITE. Insert `count` (1-1000) blank rows above or below the row named by `at` (a row number or cell reference such as `5` or `A5`; whole rows, not a range). `position` is `above` (default) or `below`.";
let patched = 0;
const walk = (o, seen = new Set()) => {
  if (!o || typeof o !== "object" || seen.has(o)) return; seen.add(o);
  if (o.name === "insert_rows" && typeof o.description === "string") { if (process.env.EVAL_VARIANT === "v1") { o.description = V1; patched++; } }
  for (const v of Object.values(o)) walk(v, seen);
};
for (const v of Object.values(defs)) walk(v);
if (process.env.EVAL_VARIANT === "v1" && !patched) { process.stderr.write("eval: insert_rows description not found to patch\n"); process.exit(3); }
await import(pathToFileURL(path.join(root, "dist", "excel_mcp", "src", "cli.js")).href);
