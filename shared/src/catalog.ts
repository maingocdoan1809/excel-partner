import { exposedToolNamesForHost, TOOL_HOSTS, type HostId, type ToolGroups } from "../../protocol/src/index.js";
import { ALL_TOOL_DEFINITIONS } from "./tool-definitions.js";
import { definitionForHost } from "./host-overrides.js";

/** Shared Office Partner catalog and external-MCP bridge tool definitions. */
export function officeToolCatalog(host: HostId, allowOfficeJs: boolean, capabilities: ToolGroups) {
  const exposed = new Set(exposedToolNamesForHost(host, allowOfficeJs, capabilities));
  return ALL_TOOL_DEFINITIONS.filter(tool => exposed.has(tool.name)).map(tool => definitionForHost(host, tool)).map(tool => ({
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
