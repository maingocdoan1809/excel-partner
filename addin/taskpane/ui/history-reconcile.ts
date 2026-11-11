import type { StoredMessage, StoredSession, StoredTool } from "../../../protocol/src/history.js";
function stable(value: unknown): string {
 return JSON.stringify(value, (_key, item) => item && typeof item === "object" && !Array.isArray(item)
   ? Object.fromEntries(Object.keys(item).sort().filter(key => item[key] !== undefined).map(key => [key, item[key]])) : item);
}
const same = (a: unknown, b: unknown) => stable(a) === stable(b);
function toolPrefix(a: StoredTool, b: StoredTool) {
 if (same(a, b)) return true;
 if (a.state !== "running") return false;
 const { state: _as, result: _ar, error: _ae, ...before } = a;
 const { state: _bs, result: _br, error: _be, ...after } = b;
 return same(before, after);
}
function messagePrefix(a: StoredMessage, b: StoredMessage) {
 if (same(a, b)) return true;
 if (a.kind !== b.kind) return false;
 if (a.kind === "user") {
   const { delivery: _ad, ...before } = a, { delivery: _bd, ...after } = b;
   return same(before, after); // delivery is UI state, not different user content
 }
 if (a.kind === "assistant") {
   const { text: _a, ...before } = a, { text: _b, ...after } = b;
   return same(before, after) && b.text.startsWith(a.text);
 }
 if (a.kind === "tool_group" || a.kind === "tool") {
   const { tools: _at, tool: _a, ...before } = a, { tools: _bt, tool: _b, ...after } = b;
   const left = a.tools ?? (a.tool ? [a.tool] : []), right = b.tools ?? (b.tool ? [b.tool] : []);
   return same(before, after) && left.length <= right.length && left.every((tool, i) => toolPrefix(tool, right[i]!));
 }
 return false;
}
/** Only monotonic transcript growth is safe to rebase automatically. */
export function historyPrefix(before: StoredSession, after: StoredSession): boolean {
 if (before.provider !== after.provider || before.providerSessionId && after.providerSessionId && before.providerSessionId !== after.providerSessionId) return false;
 return before.messages.length <= after.messages.length && before.changes.length <= after.changes.length
   && before.messages.every((message, i) => messagePrefix(message, after.messages[i]!))
   && before.changes.every((change, i) => same(change, after.changes[i]));
}
