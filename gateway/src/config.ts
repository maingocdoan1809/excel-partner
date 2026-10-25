import { z } from "zod";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { ProviderIdSchema } from "../../protocol/src/index.js";
import { ApiError } from "./protocol.js";

const RouteSchema = z.object({ provider: ProviderIdSchema, model: z.string().min(1).nullable().default(null), effort: z.enum(["low", "medium", "high"]).default("medium") });
export type Route = z.infer<typeof RouteSchema>;
export const ConfigSchema = z.object({
  port: z.number().int().min(1).max(65535).default(8766),
  defaultRoute: RouteSchema.default({ provider: "copilot", model: null, effort: "medium" }),
  aliases: z.record(RouteSchema).default({}),
  modelNames: z.record(z.string().trim().min(1).max(180)).default({}),
  enabledProviders: z.array(ProviderIdSchema).min(1).default(["copilot"]),
  allowedOrigins: z.array(z.string().url()).default(["https://pivot.claude.ai"]),
  requestTimeoutMs: z.number().int().min(1000).max(30 * 60_000).default(180_000),
  continuationTimeoutMs: z.number().int().min(1000).max(60 * 60_000).default(600_000),
  idleTimeoutMs: z.number().int().min(1000).max(24 * 60 * 60_000).default(30 * 60_000),
  maxSessions: z.number().int().min(1).max(64).default(8),
  maxBodyBytes: z.number().int().min(1024).max(64 * 1024 * 1024).default(32 * 1024 * 1024),
  maxPromptChars: z.number().int().min(1024).max(4 * 1024 * 1024).default(200_000),
  warmupOnStart: z.boolean().default(false),
  probeCacheTtlMs: z.number().int().min(1000).max(60 * 60_000).default(600_000),
  advertiseCompatibilityModels: z.boolean().default(false),
}).superRefine((config, ctx) => {
  const seen = new Set<string>();
  for (const [id, name] of Object.entries(config.modelNames)) {
    const visible = /claude|sonnet|opus|haiku|anthropic/i.test(name) ? name : `${name} (Claude API)`;
    if (seen.has(visible)) ctx.addIssue({ code: "custom", path: ["modelNames", id], message: "Model display names must be unique" });
    seen.add(visible);
  }
});
export type GatewayConfig = z.infer<typeof ConfigSchema>;
export const dataDir = () => process.env.AGENT_GATEWAY_HOME ?? path.join(process.env.LOCALAPPDATA ?? path.join(os.homedir(), ".config"), "Agent Gateway");
export function loadConfig(dir = dataDir()): GatewayConfig {
  const file = path.join(dir, "config.json");
  return ConfigSchema.parse(fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, "utf8")) : {});
}
export function routeModel(model: string, config: GatewayConfig, allowOfficeProbeFallback = false): Route {
  if (!Object.hasOwn(config.aliases, model)) {
    for (const id of Object.keys(config.modelNames)) if (publicModelId(id, config) === model) { model = id; break; }
  }
  // Claude for Excel filters model IDs to Claude/Anthropic-looking names.
  // This compatibility namespace preserves the actual provider/native ID.
  if (!Object.hasOwn(config.aliases, model) && model.startsWith("anthropic/")) return routeModel(model.slice("anthropic/".length), config);
  if (!Object.hasOwn(config.aliases, model) && model.endsWith(" [Claude API]")) return routeModel(model.slice(0, -" [Claude API]".length), config);
  let route = config.aliases[model];
  if (!route) {
    const slash = model.indexOf("/");
    if (slash > 0) {
      const parsedProvider = ProviderIdSchema.safeParse(model.slice(0, slash));
      if (!parsedProvider.success) throw new ApiError(400, "invalid_request_error", `Unknown provider: ${model.slice(0, slash)}`);
      const provider = parsedProvider.data;
      const native = model.slice(slash + 1);
      if (!native) throw new ApiError(400, "invalid_request_error", "A native model ID is required after provider/");
      route = { provider, model: native === "default" ? null : native, effort: config.defaultRoute.effort };
    } else if ((allowOfficeProbeFallback && model.startsWith("claude-"))
      || (config.advertiseCompatibilityModels && ["claude-opus-4-7", "claude-sonnet-4-6", "claude-haiku-4-5-20251001"].includes(model))) route = config.aliases[model] ?? config.defaultRoute;
    else throw new ApiError(400, "invalid_request_error", "Use provider/model or a configured model alias");
  }
  if (!config.enabledProviders.includes(route.provider)) throw new ApiError(403, "permission_error", `Provider '${route.provider}' is not enabled`);
  return { ...route };
}
/** Office uses model IDs as labels, so a named public alias doubles as its label. */
export function publicModelId(id: string, config: GatewayConfig): string {
  const name = config.modelNames[id];
  if (!name) return id;
  return /claude|sonnet|opus|haiku|anthropic/i.test(name) ? name : `${name} (Claude API)`;
}
