import { z } from "zod";
import { TokenUsageSchema } from "./usage.js";

/** Why a turn ended. */
export const TurnFinishReasonSchema = z.enum(["completed", "cancelled", "error"]);

/**
 * Provider output normalised by every adapter so the task pane never has to
 * understand a CLI-specific stream format.
 */
export const ProviderEventSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("followup_status"),
    requestId: z.string(),
    id: z.string(),
    state: z.enum(["pending", "delivered", "next_turn", "rejected"]),
    message: z.string().optional(),
  }),
  z.object({
    type: z.literal("session_info"),
    requestId: z.string(),
    provider: z.enum(["claude", "copilot", "codex", "opencode"]),
    providerSessionId: z.string().min(1).max(128),
    title: z.string().max(200).optional(),
  }),
  z.object({
    type: z.literal("usage_update"),
    requestId: z.string(),
    /** Cumulative snapshot for this turn, replacing earlier snapshots. */
    usage: TokenUsageSchema,
  }),
  z.object({
    /** The provider CLI is being opened for this request; `turn_started` follows once it is ready. */
    type: z.literal("session_starting"),
    requestId: z.string(),
  }),
  z.object({
    type: z.literal("provider_progress"),
    requestId: z.string(),
    state: z.enum(["waiting", "retrying"]),
    message: z.string().max(300),
  }),
  z.object({
    type: z.literal("turn_started"),
    requestId: z.string(),
  }),
  z.object({
    /** What the model reasons before answering (full text or the provider's summary of it), shown while the turn runs. */
    type: z.literal("thinking_delta"),
    requestId: z.string(),
    text: z.string(),
  }),
  z.object({
    type: z.literal("message_delta"),
    requestId: z.string(),
    text: z.string(),
  }),
  z.object({
    type: z.literal("message_completed"),
    requestId: z.string(),
    text: z.string(),
  }),
  z.object({
    type: z.literal("tool_started"),
    requestId: z.string(),
    callId: z.string(),
    tool: z.string(),
    args: z.unknown(),
  }),
  z.object({
    type: z.literal("tool_finished"),
    requestId: z.string(),
    callId: z.string(),
    tool: z.string(),
    ok: z.boolean(),
    error: z.string().nullable(),
  }),
  z.object({
    type: z.literal("turn_finished"),
    requestId: z.string(),
    reason: TurnFinishReasonSchema,
  }),
  z.object({
    type: z.literal("error"),
    requestId: z.string().nullable(),
    code: z.string(),
    message: z.string(),
  }),
]);
export type ProviderEvent = z.infer<typeof ProviderEventSchema>;
