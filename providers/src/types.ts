import type {
  ProviderConnectionStatus,
  ProviderDetection,
  ProviderEvent,
  ProviderId,
  SelectionSnapshot,
} from "../../protocol/src/index.js";

/**
 * Everything an adapter needs to launch its CLI wired to the local Excel MCP
 * server. The CLI spawns the MCP process; that process calls back into the
 * bridge over `bridgeUrl` using `token` to reach the right workbook session.
 */
export interface McpEndpoint {
  bridgeUrl: string;
  sessionId: string;
  token: string;
  command: string;
  args: string[];
  /**
   * Environment the MCP process needs to call back into the bridge
   * (EXCEL_BRIDGE_URL, EXCEL_SESSION_ID, EXCEL_SESSION_TOKEN and, for the
   * self-signed localhost certificate, NODE_EXTRA_CA_CERTS).
   */
  env: Record<string, string>;
}

/** Name the Excel MCP server is registered under in every CLI. */
export const MCP_SERVER_NAME = "excel";

export interface StartSessionContext {
  /** Present native provider questions in the pane and wait for the person's answers. */
  requestUserInput?: (request: import("../../protocol/src/index.js").UserInputRequest, signal?: AbortSignal) => Promise<import("../../protocol/src/index.js").UserInputResponse>;
  /** Ask the connected pane; its current Ask/Auto setting decides approval. */
  authorizeTool?: (args: unknown) => Promise<boolean>;
  systemPrompt?: string;
  workbookConnectionId: string;
  sessionId: string;
  model?: string | null;
  effort?: "low" | "medium" | "high";
  resumeSessionId?: string;
  mcp: McpEndpoint;
  /** Aborted when the workbook/task pane disconnects. */
  signal: AbortSignal;
}

export interface SendTurnInput {
  systemPrompt?: string;
  providerSessionId: string;
  requestId: string;
  model?: string | null;
  effort?: "low" | "medium" | "high";
  prompt: string;
  /** Original user text before bridge context and skill expansion. */
  userPrompt?: string;
  selection: SelectionSnapshot | null;
  /** Aborted when the user hits stop or the pane disconnects. */
  signal: AbortSignal;
}

export interface ProviderModel {
  id: string;
  name: string;
}

export interface ProviderModelCatalog {
  models: ProviderModel[];
  defaultModelId: string | null;
}

/**
 * The single contract every CLI provider implements. The bridge only ever sees
 * normalised `ProviderEvent`s, never CLI-specific output.
 */
export interface ProviderAdapter {
  readonly id: ProviderId;
  /** Whether the CLI is installed and what it supports. */
  detect(force?: boolean): Promise<ProviderDetection>;
  /** Models advertised by the installed CLI, when it offers a catalog. */
  listModels?(): Promise<ProviderModelCatalog>;
  /** Read the native title without a model request. Missing titles stay unknown. */
  sessionTitle?(nativeSessionId: string): Promise<string | undefined>;
  /** Whether the CLI is logged in and ready to accept a turn. */
  connectionStatus(): Promise<ProviderConnectionStatus>;
  /** Launch the CLI/session and return a provider-owned session id. */
  startSession(ctx: StartSessionContext): Promise<string>;
  /** Run one user turn, streaming normalised events. */
  sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent>;
  /** Stop the running turn without closing the session. */
  cancelTurn(providerSessionId: string): Promise<void>;
  /** End the session and release the CLI process. */
  closeSession(providerSessionId: string): Promise<void>;
  /** Release all resources; called on bridge shutdown. */
  dispose(): Promise<void>;
}

export interface ProviderRuntime {
  bridgeUrl: string;
  mcp: { command: string; args: string[] };
  logger: { debug(msg: string, extra?: unknown): void; warn(msg: string, extra?: unknown): void };
}
