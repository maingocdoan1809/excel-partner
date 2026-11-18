// Per-variant summary of .claude/hillclimb/excel-tools/<variant>/results.jsonl, plus sanity checks on every row.
import fs from "node:fs";
const dir = ".claude/hillclimb/excel-tools";
const load = v => fs.readFileSync(`${dir}/${v}/results.jsonl`, "utf8").split("\n").filter(Boolean).map(l => JSON.parse(l));
const variants = process.argv.slice(2).length ? process.argv.slice(2) : ["baseline"];
const pct = (a, k) => (a.reduce((t, r) => t + r.grade[k], 0) / a.length * 100).toFixed(0) + "%";
for (const v of variants) {
  const rows = load(v);
  const bad = rows.filter(r => !r.usage.output_tokens || !r.latency_s || r.latency_s <= 0);
  console.log(`\n== ${v}: ${rows.length} rows | pass ${pct(rows, "pass")} tool ${pct(rows, "tool")} state ${pct(rows, "state")} reply ${pct(rows, "reply")} | rows with 0 tokens or 0 seconds: ${bad.length}`);
  const lat = rows.map(r => r.latency_s).sort((a, b) => a - b); const calls = rows.map(r => r.tool_calls).sort((a, b) => a - b);
  console.log(`   latency s  min/median/max ${lat[0]}/${lat[lat.length >> 1]}/${lat.at(-1)} | tool calls min/median/max ${calls[0]}/${calls[calls.length >> 1]}/${calls.at(-1)} | unsimulated tool calls: ${rows.filter(r => r.meta.unsimulated_tools.length).length} rows`);
  for (const tag of [...new Set(rows.map(r => r.tags[0]))]) { const a = rows.filter(r => r.tags[0] === tag); console.log(`   ${tag.padEnd(9)} n=${String(a.length).padStart(2)}  pass ${pct(a, "pass").padStart(4)}  tool ${pct(a, "tool").padStart(4)}  state ${pct(a, "state").padStart(4)}  reply ${pct(a, "reply").padStart(4)}`); }
  console.log("   cases with a failure:");
  for (const id of [...new Set(rows.map(r => r.prompt_id))]) { const a = rows.filter(r => r.prompt_id === id); const f = a.filter(r => !r.grade.pass); if (f.length) console.log(`     ${id} ${a.length - f.length}/${a.length} pass | failed checks: ${f.map(r => ["tool", "state", "reply"].filter(k => !r.grade[k]).join("+")).join(", ")} | ${f[0].explanation.pass.slice(0, 150)}`); }
}
