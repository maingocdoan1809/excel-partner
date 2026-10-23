import { ProviderCommandSchema, type ProviderCommand } from "../../protocol/src/commands.js";

/** Preserve provider names; the pane owns controls that must stay in sync with its settings. */
const CLIENT_ACTIONS: Record<string, ProviderCommand["action"]> = {
  clear: "new", new: "new", resume: "history", sessions: "history", model: "model", models: "model",
  effort: "effort", permissions: "permissions", login: "login", theme: "theme", copy: "copy", export: "export", help: "help",
};

export function normalizeCommands(items: unknown): ProviderCommand[] {
  if (!Array.isArray(items)) return [];
  const seen = new Set<string>();
  return items.flatMap(item => {
    if (!item || typeof item.name !== "string") return [];
    const name = item.name.replace(/^\//, "");
    const action = CLIENT_ACTIONS[name];
    const parsed = ProviderCommandSchema.safeParse({ name, description: item.description ?? "",
      source: item.kind === "skill" || item.source === "skill" || item.builtin === false ? "skill" : "builtin",
      execution: action ? "client" : item.execution ?? "native", action: action ?? item.action,
      inputHint: item.input?.hint ?? item.argumentHint ?? item.inputHint, reason: item.reason });
    if (!parsed.success || seen.has(name)) return [];
    seen.add(name);
    return [parsed.data];
  });
}

/** Coalesce concurrent menu requests. Failed discovery can be retried immediately. */
export function cachedCommands(load: () => Promise<ProviderCommand[]>, ttlMs = 30_000) {
  let pending: Promise<ProviderCommand[]> | undefined;
  let expires = 0;
  return () => {
    if (!pending || Date.now() >= expires) {
      expires = Date.now() + ttlMs;
      pending = load().catch(error => { pending = undefined; expires = 0; throw error; });
    }
    return pending;
  };
}
