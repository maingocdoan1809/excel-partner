/**
 * A plain line diff for showing a person what changed between two versions of a script. Scripts are small
 * (a few hundred lines), so a straightforward longest-common-subsequence table is enough.
 */
const MAX_CELLS = 4_000_000;

export function lineDiff(before: string, after: string, context = 2): string[] {
  const a = before.split("\n"), b = after.split("\n");
  if (a.length * b.length > MAX_CELLS) return [`(the file is too large to compare line by line: ${a.length} lines before, ${b.length} after)`];
  const columns = b.length + 1;
  const table = new Int32Array((a.length + 1) * columns);
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      table[i * columns + j] = a[i] === b[j] ? table[(i + 1) * columns + j + 1]! + 1 : Math.max(table[(i + 1) * columns + j]!, table[i * columns + j + 1]!);
    }
  }
  const lines: Array<{ mark: " " | "-" | "+"; text: string }> = [];
  let i = 0, j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { lines.push({ mark: " ", text: a[i]! }); i++; j++; }
    else if (table[(i + 1) * columns + j]! >= table[i * columns + j + 1]!) lines.push({ mark: "-", text: a[i++]! });
    else lines.push({ mark: "+", text: b[j++]! });
  }
  while (i < a.length) lines.push({ mark: "-", text: a[i++]! });
  while (j < b.length) lines.push({ mark: "+", text: b[j++]! });

  const keep = new Array<boolean>(lines.length).fill(false);
  lines.forEach((line, index) => {
    if (line.mark === " ") return;
    for (let k = Math.max(0, index - context); k <= Math.min(lines.length - 1, index + context); k++) keep[k] = true;
  });
  const out: string[] = [];
  let gap = false;
  lines.forEach((line, index) => {
    if (!keep[index]) { gap = out.length > 0; return; }
    if (gap) { out.push("  ..."); gap = false; }
    out.push(`${line.mark} ${line.text}`);
  });
  return out;
}
