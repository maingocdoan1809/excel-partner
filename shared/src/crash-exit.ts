/** Exit code of a bridge or gateway that stopped on an unexpected error; the supervisor restarts it. */
export const CRASH_EXIT_CODE = 70;

/**
 * What is logged about an unexpected error: its type, code and stack frames. The message is left out because it can quote
 * a prompt, document content or a provider's output.
 */
export function crashSummary(error: unknown): { name: string; code?: string; frames?: string[] } {
  if (!(error instanceof Error)) return { name: typeof error };
  const code = (error as NodeJS.ErrnoException).code;
  const frames = error.stack?.split("\n").filter((line) => /^\s+at /.test(line)).slice(0, 12).map((line) => line.trim());
  return { name: error.name, ...(typeof code === "string" ? { code } : {}), ...(frames?.length ? { frames } : {}) };
}
