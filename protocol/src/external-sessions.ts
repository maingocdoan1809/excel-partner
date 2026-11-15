import { z } from "zod";
import type { StoredMessage } from "./history.js";

/** A conversation a provider's own tools (its terminal CLI, its VS Code extension) saved on this computer. */
export const ExternalSessionSummarySchema = z.object({
  /** The provider's own id for the conversation; the pane passes it back to import it. */
  id: z.string().min(1).max(256),
  title: z.string().max(1000),
  /** Where it was held, for people: "Terminal", "VS Code", … */
  origin: z.string().max(64),
  /** The folder it was held in, when the provider recorded one. */
  cwd: z.string().max(1024).optional(),
  updatedAt: z.number().finite(),
  messageCount: z.number().int().nonnegative().optional(),
});
export type ExternalSessionSummary = z.infer<typeof ExternalSessionSummarySchema>;

/** One of them read in full and made resumable: `providerSessionId` opens it with `resumeSessionId`. */
export interface ImportedSession {
  providerSessionId: string;
  title: string;
  startedAt: number;
  updatedAt: number;
  messages: StoredMessage[];
}
