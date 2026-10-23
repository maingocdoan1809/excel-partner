/** One entry of the `/` menu: a skill from the bridge or a built-in command. */
export interface SlashCommand {
  name: string;
  description: string;
  kind: "skill" | "builtin" | "provider";
  command?: import("../../../protocol/src/commands.js").ProviderCommand;
}

export const BUILTIN_COMMANDS: SlashCommand[] = [{ name: "clear", description: "Clear this conversation", kind: "builtin" }];

const MAX_ITEMS = 200;

/**
 * The text typed after the leading `/` while the user is still on the command word
 * (`/ast` -> "ast", `/` -> ""); null once there is a space or the text does not start with `/`.
 */
export function slashQuery(text: string): string | null {
  const match = /^\/([\w:-]*)$/.exec(text);
  return match ? match[1]!.toLowerCase() : null;
}

/** Prefix matches on the name first, then substring matches on name or description. */
export function filterCommands(commands: SlashCommand[], query: string): SlashCommand[] {
  if (!query) return commands.slice(0, MAX_ITEMS);
  const rank = (c: SlashCommand): number => {
    const name = c.name.toLowerCase();
    if (name.startsWith(query)) return 0;
    if (name.includes(query)) return 1;
    if (c.description.toLowerCase().includes(query)) return 2;
    return 3;
  };
  return commands
    .map((c) => ({ c, r: rank(c) }))
    .filter((x) => x.r < 3)
    .sort((a, b) => a.r - b.r || a.c.name.localeCompare(b.c.name))
    .slice(0, MAX_ITEMS)
    .map((x) => x.c);
}

/** `/name rest of the message` -> the command word, for highlighting in the chat. */
export function leadingCommand(text: string): { command: string; rest: string } | null {
  const match = /^(\/[\w:-]+)(\s[\s\S]*)?$/.exec(text);
  return match ? { command: match[1]!, rest: match[2] ?? "" } : null;
}

/** The completed leading command to display and edit as a badge. */
export function slashBadgeRange(text: string, selectedToken: string | null): { start: number; end: number } | null {
  const lead = leadingCommand(text);
  if (!lead || (!lead.rest && lead.command !== selectedToken && lead.command !== "/clear")) return null;
  return { start: 0, end: lead.command.length };
}
