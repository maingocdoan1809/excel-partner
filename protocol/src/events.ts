import { z } from "zod";

/** Why a turn ended. */
export const TurnFinishReasonSchema = z.enum(["completed", "cancelled", "error"]);
export type TurnFinishReason = z.infer<typeof TurnFinishReasonSchema>;

/**
 * Provider output normalised by every adapter so the task pane never has to
 * understand a CLI-specific stream format.
 */
export const ProviderEventSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("turn_started"),
    requestId: z.string(),
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

export type ProviderEventType = ProviderEvent["type"];
