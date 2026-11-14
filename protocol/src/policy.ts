import { z } from "zod";
import { ProviderIdSchema } from "./providers.js";

/**
 * Machine-wide admin policy, read from `%ProgramData%\Office Partner\policy.json` (or `OFFICE_PARTNER_POLICY_FILE`).
 * Only administrators can write there, so users cannot loosen it. A missing file means no restrictions.
 * A file that exists but does not parse fails closed: see `LOCKED_DOWN_POLICY`.
 */
export const AdminPolicySchema = z.object({
  /** Providers hidden from the pane and refused by the bridge. */
  disabledProviders: z.array(ProviderIdSchema).default([]),
  /** May the pane's Auto mode approve tools without asking? Off forces Ask. */
  allowAutoApprove: z.boolean().default(true),
  /** run_python and skill scripts (code that runs as the user). */
  allowCodeExecution: z.boolean().default(true),
  /** User-configured external MCP servers. */
  allowExternalMcp: z.boolean().default(true),
  /** Provider web search / web fetch tools. */
  allowWebSearch: z.boolean().default(true),
  /** Browser automation (Playwright) and saved browser sign-ins. */
  allowBrowserAutomation: z.boolean().default(true),
  /** execute_office_js (arbitrary Office.js in the pane). */
  allowOfficeJs: z.boolean().default(true),
  /** May the product update itself (pane Update button and `office-partner update`)? */
  autoUpdate: z.boolean().default(true),
  /** Exact version every update must install, e.g. "0.1.23". Updates to any other version are refused. */
  pinVersion: z.string().regex(/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/).optional(),
  /** npm registry used for update lookups and installs; defaults to npm's own configuration. */
  registry: z.string().url().optional(),
  /** Delete chat history older than this many days (0 or absent keeps everything). */
  historyRetentionDays: z.number().int().min(0).optional(),
}).strict();
export type AdminPolicy = z.infer<typeof AdminPolicySchema>;

export const OPEN_POLICY: AdminPolicy = AdminPolicySchema.parse({});

/** What applies when a policy file exists but is unreadable or invalid: every risky feature off, no self-update. */
export const LOCKED_DOWN_POLICY: AdminPolicy = {
  ...OPEN_POLICY,
  allowAutoApprove: false, allowCodeExecution: false, allowExternalMcp: false, allowWebSearch: false,
  allowBrowserAutomation: false, allowOfficeJs: false, autoUpdate: false,
};

/** Parses policy text. Returns the locked-down policy (and the reason) when the text is not a valid policy. */
export function parseAdminPolicy(text: string): { policy: AdminPolicy; error?: string } {
  let raw: unknown;
  try { raw = JSON.parse(text); } catch (error) { return { policy: LOCKED_DOWN_POLICY, error: `policy.json is not valid JSON: ${(error as Error).message}` }; }
  const parsed = AdminPolicySchema.safeParse(raw);
  return parsed.success ? { policy: parsed.data } : { policy: LOCKED_DOWN_POLICY, error: `policy.json is invalid: ${parsed.error.issues.map(i => `${i.path.join(".") || "(root)"} ${i.message}`).join("; ")}` };
}
