import { z } from "zod";

export const ProviderCommandSchema = z.object({
  name: z.string().regex(/^[\w:-]+$/).max(200),
  description: z.string().max(8000),
  source: z.enum(["builtin", "skill"]).default("builtin"),
  execution: z.enum(["native", "client", "unsupported"]),
  inputHint: z.string().max(2000).optional(),
  action: z.enum(["new", "history", "model", "effort", "permissions", "usage", "login", "theme", "copy", "export", "help", "status"]).optional(),
  reason: z.string().optional(),
});
export type ProviderCommand = z.infer<typeof ProviderCommandSchema>;
export interface NativeCommand { name: string; args: string }

export function parseNativeCommand(text: string): NativeCommand | null {
  const match = /^\/([\w:-]+)(?:\s+([\s\S]*))?$/.exec(text.trim());
  return match ? { name: match[1]!, args: match[2]?.trim() ?? "" } : null;
}
