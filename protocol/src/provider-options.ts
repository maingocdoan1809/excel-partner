import { z } from "zod";

/**
 * A setting the CLI itself offers for the running session (a mode such as plan or build, for example), as the CLI describes it.
 * The pane draws what it is given and passes the choice back; it knows nothing about what an option does.
 */
export const ProviderOptionSchema = z.object({
  id: z.string().min(1).max(100),
  label: z.string().min(1).max(100),
  /** The choice in force now. */
  value: z.string().max(200),
  choices: z.array(z.object({ value: z.string().max(200), label: z.string().min(1).max(100) })).min(1).max(40),
});
export type ProviderOption = z.infer<typeof ProviderOptionSchema>;
