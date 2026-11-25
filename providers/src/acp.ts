import { MCP_SERVER_NAME } from "./types.js";
import type { ProviderModelCatalog, StartSessionContext } from "./types.js";
import type { AcpConfigOption } from "./options.js";

/**
 * Helpers shared by the ACP adapters. They speak
 * the same protocol over stdio, so tool naming and permission answers are
 * identical.
 */

/** Extract the bare Excel tool name from an ACP tool_call (`excel-get_selection`, `excel/get_selection`, ...). */
export function toolNameOf(update: Record<string, any>): string;
export function toolNameOf(update: Record<string, any>, detail: true): { name: string; isExcel: boolean };
export function toolNameOf(update: Record<string, any>, detail?: true): string | { name: string; isExcel: boolean } {
  const candidates = [update._meta?.toolName, update.rawInput?.toolName, update.title, update.kind]
    .filter((v): v is string => typeof v === "string");
  let name = candidates[0] ?? "unknown";
  let isExcel = false;
  for (const candidate of candidates) {
    const match = new RegExp(`${MCP_SERVER_NAME}[-_/:.]+([a-z_]+)`, "i").exec(candidate);
    if (match) {
      name = match[1]!;
      isExcel = true;
      break;
    }
  }
  return detail ? { name, isExcel } : name;
}

/** Flatten the content of a finished tool call into plain text. */
export function contentText(update: Record<string, any>): string {
  const parts: string[] = [];
  for (const item of update.content ?? []) {
    const text = item?.content?.text ?? item?.text;
    if (typeof text === "string") parts.push(text);
  }
  if (parts.length === 0 && update.rawOutput) {
    parts.push(typeof update.rawOutput === "string" ? update.rawOutput : JSON.stringify(update.rawOutput));
  }
  return parts.join(" ");
}

/**
 * Answer an ACP `session/request_permission` request. Excel MCP tools are
 * allowed, everything else (shell, file edits, other MCP servers) is
 * rejected, so a provider can only ever touch the workbook.
 */
export function acpPermissionOutcome(
  params: unknown,
  allowedNativeTools: readonly string[] = [],
): { outcome: { outcome: "selected"; optionId: string } } | { outcome: { outcome: "cancelled" } } {
  const request = (params ?? {}) as {
    toolCall?: Record<string, any>;
    options?: Array<{ optionId: string; kind?: string }>;
  };
  const options = request.options ?? [];
  const isExcel = toolNameOf(request.toolCall ?? {}, true).isExcel;
  const call = request.toolCall ?? {};
  const native = [call._meta?.toolName, call.rawInput?.toolName, call.title].some(name => typeof name === "string" && allowedNativeTools.includes(name));
  const allowed = isExcel || native;
  const wanted = allowed ? ["allow_once", "allow_always"] : ["reject_once", "reject_always"];
  const choice = options.find((option) => wanted.includes(option.kind ?? ""))
    ?? options.find((option) => (allowed ? /allow/i : /reject/i).test(`${option.kind ?? ""} ${option.optionId ?? ""}`));
  if (choice) return { outcome: { outcome: "selected", optionId: choice.optionId } };
  return { outcome: { outcome: "cancelled" } };
}

/** Non-Office permissions follow the pane's Ask/Auto flow. Office tools are gated by the bridge itself. */
export async function authorizeAcpRequest(method: string, params: unknown, authorize?: StartSessionContext["authorizeTool"]): Promise<unknown> {
  if (method !== "session/request_permission") throw new Error(`unsupported client method: ${method}`);
  const request = (params ?? {}) as { toolCall?: Record<string, any>; options?: Array<{ optionId: string; kind?: string }> };
  if (toolNameOf(request.toolCall ?? {}, true).isExcel) return acpPermissionOutcome(params);
  let approved = false;
  try { approved = await authorize?.(request.toolCall ?? {}) ?? false; } catch { /* fail closed */ }
  const choice = request.options?.find(option => option.kind === (approved ? "allow_once" : "reject_once"));
  return choice ? { outcome: { outcome: "selected", optionId: choice.optionId } } : { outcome: { outcome: "cancelled" } };
}

/** The config option that selects the model, if the agent advertises one. */
export function acpModelOption(options?: AcpConfigOption[]) {
  return options?.find(item => item.category === "model" || item.id === "model");
}

/** ACP agents can advertise models in config options or the original models field. */
export function acpModels(session: { configOptions?: AcpConfigOption[]; models?: {
  currentModelId?: string; availableModels?: Array<{ modelId: string; name?: string }>;
} }): ProviderModelCatalog {
  const option = acpModelOption(session.configOptions);
  const entries = option?.options?.flatMap(item => item.options ?? (item.value ? [{ value: item.value, name: item.name }] : []));
  const seen = new Set<string>();
  const models = (entries ?? session.models?.availableModels?.map(item => ({ value: item.modelId, name: item.name })) ?? [])
    .flatMap(item => {
      if (!item.value || seen.has(item.value)) return [];
      seen.add(item.value);
      return [{ id: item.value, name: item.name || item.value }];
    });
  return { models, defaultModelId: typeof option?.currentValue === "string" ? option.currentValue : session.models?.currentModelId ?? null };
}
