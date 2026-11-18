// Paired comparison of two variants: per-case mean pass over reps, mean difference and a Wald 95% interval over the per-case differences.
import fs from "node:fs";
const dir = ".claude/hillclimb/excel-tools"; const [a, b] = process.argv.slice(2).length ? process.argv.slice(2) : ["baseline", "v1"];
const load = v => fs.readFileSync(`${dir}/${v}/results.jsonl`, "utf8").split("\n").filter(Boolean).map(l => JSON.parse(l));
const A = load(a), B = load(b);
const byCase = rows => { const m = new Map(); for (const r of rows) { const e = m.get(r.prompt_id) ?? { tag: r.tags[0], v: [] }; e.v.push(r.grade.pass); m.set(r.prompt_id, e); } return m; };
const ma = byCase(A), mb = byCase(B); const mean = x => x.reduce((s, y) => s + y, 0) / x.length;
const ci = d => { const m = mean(d), sd = Math.sqrt(d.reduce((s, x) => s + (x - m) ** 2, 0) / Math.max(1, d.length - 1)), h = 1.96 * sd / Math.sqrt(d.length); return `${(m * 100).toFixed(1)} pts  [${((m - h) * 100).toFixed(1)}, ${((m + h) * 100).toFixed(1)}]  n=${d.length}`; };
const deltas = (filter = () => true) => [...ma].filter(([id, e]) => mb.has(id) && filter(e)).map(([id, e]) => ({ id, tag: e.tag, d: mean(mb.get(id).v) - mean(e.v), a: mean(e.v), b: mean(mb.get(id).v) }));
const all = deltas(); console.log(`${b} minus ${a}, per-case paired, all cases: ${ci(all.map(x => x.d))}`);
for (const tag of [...new Set(all.map(x => x.tag))]) { const d = all.filter(x => x.tag === tag); console.log(`  ${tag.padEnd(9)} ${ci(d.map(x => x.d))}`); }
console.log("cases whose pass rate differs:"); for (const x of all.filter(x => Math.abs(x.d) > 1e-9)) console.log(`  ${x.id} ${a} ${(x.a * 100).toFixed(0)}% -> ${b} ${(x.b * 100).toFixed(0)}%`);
const perf = rows => ({ calls: mean(rows.map(r => r.tool_calls)), lat: mean(rows.map(r => r.latency_s)), out: mean(rows.map(r => r.usage.output_tokens)) });
const pa = perf(A), pb = perf(B); console.log(`mean per run  ${a}: ${pa.calls.toFixed(1)} calls, ${pa.lat.toFixed(1)} s, ${pa.out.toFixed(0)} out tokens | ${b}: ${pb.calls.toFixed(1)} calls, ${pb.lat.toFixed(1)} s, ${pb.out.toFixed(0)} out tokens`);
