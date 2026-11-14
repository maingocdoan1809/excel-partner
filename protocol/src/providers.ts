import { z } from "zod";

/** CLI providers supported by the bridge. */
export const ProviderIdSchema = z.enum(["claude", "copilot", "codex", "opencode"]);
export type ProviderId = z.infer<typeof ProviderIdSchema>;

export const PROVIDER_IDS = ProviderIdSchema.options;

/** The provider a new install, a new chat or a request that names none starts with. */
export const DEFAULT_PROVIDER: ProviderId = "copilot";

export const ProviderConnectionStatusSchema = z.enum([
  "unknown",
  "not_installed",
  "login_required",
  "ready",
  "error",
]);
export type ProviderConnectionStatus = z.infer<typeof ProviderConnectionStatusSchema>;

/** How the pane signs in: a code to enter on a page, a page that hands back a code, or a list of accounts to pick from. */
export const ProviderLoginKindSchema = z.enum(["device", "browser", "account"]);
export type ProviderLoginKind = z.infer<typeof ProviderLoginKindSchema>;

export const ProviderCapabilitiesSchema = z.object({
  streaming: z.boolean(),
  cancel: z.boolean(),
  sessionResume: z.boolean(),
  mcp: z.boolean(),
  acp: z.boolean().optional(),
  /** The provider has its own web search; the add-in adds no search service. */
  nativeWebSearch: z.boolean(),
  /** The provider filters MCP tools by name, so the bridge lists the names it wants exposed. */
  explicitToolNames: z.boolean(),
  /** Usage is billed in credits, which the pane shows. */
  credits: z.boolean(),
  loginKind: ProviderLoginKindSchema,
  /** The provider's CLI can sign out, so the pane offers it. */
  logout: z.boolean(),
  /** The CLI changes its model (`liveModel`) or reasoning effort (`liveEffort`) in the middle of a session, from the next turn on, so the bridge does not restart it for that. */
  liveModel: z.boolean(),
  liveEffort: z.boolean(),
});
export type ProviderCapabilities = z.infer<typeof ProviderCapabilitiesSchema>;

export const ProviderDetectionSchema = z.object({
  provider: ProviderIdSchema,
  installed: z.boolean(),
  /** Raw `--version` string when the CLI could be probed. */
  version: z.string().nullable(),
  status: ProviderConnectionStatusSchema,
  /** Human readable hint shown in the task pane when status is not `ready`. */
  hint: z.string().nullable(),
  capabilities: ProviderCapabilitiesSchema,
});
export type ProviderDetection = z.infer<typeof ProviderDetectionSchema>;
