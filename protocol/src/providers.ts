import { z } from "zod";

/** CLI providers supported by the bridge. */
export const ProviderIdSchema = z.enum(["claude", "copilot", "codex", "opencode"]);
export type ProviderId = z.infer<typeof ProviderIdSchema>;

export const PROVIDER_IDS = ProviderIdSchema.options;

export const ProviderConnectionStatusSchema = z.enum([
  "unknown",
  "not_installed",
  "login_required",
  "ready",
  "error",
]);
export type ProviderConnectionStatus = z.infer<typeof ProviderConnectionStatusSchema>;

export const ProviderCapabilitiesSchema = z.object({
  streaming: z.boolean(),
  cancel: z.boolean(),
  sessionResume: z.boolean(),
  mcp: z.boolean(),
  acp: z.boolean().optional(),
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
