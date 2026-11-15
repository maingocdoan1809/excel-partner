import { exposedToolNamesForHost, TOOL_HOSTS, type HostId, type ToolGroups } from "../../protocol/src/index.js";
import { ALL_TOOL_DEFINITIONS } from "./tool-definitions.js";
import { definitionForHost } from "./host-overrides.js";
import { GET_TOOL_DEFINITION, needsLookup, stubOf, suggestToolNames } from "./lazy-tools.js";

/**
 * Shared Office Partner catalog and external-MCP bridge tool definitions. With `capabilities.lazySchemas` every tool except
 * the host's few hot ones is offered as a stub (summary, open schema); see lazy-tools.ts.
 */
export function officeToolCatalog(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups) {
  return catalogOf(host, allowOfficeJs, capabilities, capabilities.lazySchemas === true);
}

/** The same tools with their real definitions, for `get_tool_definition` to answer from (it describes itself too). */
export function describableTools(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups) {
  return catalogOf(host, allowOfficeJs, { ...capabilities, lazySchemas: true }, false);
}

function catalogOf(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups, stubs: boolean) {
  const exposed = new Set(exposedToolNamesForHost(host, allowOfficeJs, capabilities));
  return ALL_TOOL_DEFINITIONS.filter(tool => exposed.has(tool.name)).map(tool => definitionForHost(host, tool))
    .map(tool => stubs ? stubOf(tool) : tool).map(tool => ({
      name: tool.name, description: tool.description, inputSchema: tool.inputSchema as Record<string, unknown>, annotations: tool.annotations,
    }));
}

/** Infer only from exclusive Office tools. Shared escape hatches alone do not identify a host. */
export function inferOfficeHost(names: string[]): HostId | undefined {
  const hosts = new Set<HostId>();
  for (const name of names) {
    const supported = TOOL_HOSTS[name];
    if (supported?.length === 1) hosts.add(supported[0]!);
  }
  return hosts.size === 1 ? [...hosts][0] : undefined;
}

/**
 * Characters of tool definitions (name, description, input schema) the model is offered for a host: what every request pays
 * before the first word. The usage ledger records it per turn and the catalog budget test holds it down.
 */
export function toolCatalogChars(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups): number {
  return officeToolCatalog(host, allowOfficeJs, capabilities)
    .reduce((total, tool) => total + JSON.stringify({ name: tool.name, description: tool.description, inputSchema: tool.inputSchema }).length, 0);
}

/**
 * What to add to an `invalid_arguments` error so one retry fixes it: in lazy mode the schema of a tool the model was only
 * given a summary of. Empty when the model already had the schema (full mode, or a hot tool) or the tool is unknown.
 */
export function schemaHintFor(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups, name: string): string {
  if (capabilities.lazySchemas !== true) return "";
  const tool = describableTools(host, allowOfficeJs, capabilities).find(candidate => candidate.name === name);
  return tool && needsLookup(tool) ? ` Input schema of ${name}: ${JSON.stringify(tool.inputSchema)}` : "";
}

/** What to add to an `unknown_tool` error: the closest tool names, or where to find the list. */
export function unknownToolHint(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups, name: string): string {
  const names = officeToolCatalog(host, allowOfficeJs, capabilities).map(tool => tool.name);
  const close = suggestToolNames(name, names);
  if (close.length) return ` Did you mean: ${close.join(", ")}?`;
  return capabilities.lazySchemas === true ? ` Call ${GET_TOOL_DEFINITION} with no names to list every tool.` : "";
}
