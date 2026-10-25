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
  /** Explicit client-supplied MCP tool names; omitted for the Excel tool catalog. */
  toolNames?: string[];
  /**
   * Environment the MCP process needs to call back into the bridge
   * (EXCEL_BRIDGE_URL, EXCEL_SESSION_ID, EXCEL_SESSION_TOKEN and, for the
   * self-signed localhost certificate, NODE_EXTRA_CA_CERTS).
   */
  env: Record<string, string>;
}

/** Name the Excel MCP server is registered under in every CLI. */
export const MCP_SERVER_NAME = "excel";

/** Metadata only: never include upstream response bodies, headers or prompts. */
export interface ProviderDiagnostic {
  phase: "model_retry_scheduled" | "model_retry_started";
  attempt?: number;
  status?: number;
  errorType?: string;
  nextRetryAt?: string;
}

export interface StartSessionContext {
  /** Optional owner-side diagnostics; workbook sessions need not subscribe. */
  onDiagnostic?: (event: ProviderDiagnostic) => void;
  /** Notify the owner when this runtime can no longer accept turns, including
   * forced cancellation and idle exits. This does not end the native conversation.
   */
  onRuntimeClosed?: () => void;
  /** Present native provider questions in the pane and wait for the person's answers. */
  requestUserInput?: (request: import("../../protocol/src/index.js").UserInputRequest, signal?: AbortSignal) => Promise<import("../../protocol/src/index.js").UserInputResponse>;
  /** Ask the connected pane; its current Ask/Auto setting decides approval. */
  authorizeTool?: (args: unknown) => Promise<boolean>;
  systemPrompt?: string;
  /** API gateways can replace Copilot's default agent prompt; Excel keeps append. */
  systemPromptMode?: "append" | "replace";
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
  command?: import("../../protocol/src/commands.js").NativeCommand;
  systemPrompt?: string;
  providerSessionId: string;
  requestId: string;
  model?: string | null;
  effort?: "low" | "medium" | "high";
  prompt: string;
  /** Images attached to the current API turn; bytes are kept out of prompt text and logs. */
  images?: Array<{ mediaType: "image/jpeg" | "image/png" | "image/gif" | "image/webp"; data: string }>;
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
  listCommands?(providerSessionId?: string): Promise<import("../../protocol/src/commands.js").ProviderCommand[]>;
  readonly id: ProviderId;
  /** Whether the CLI is installed and what it supports. */
  detect(force?: boolean): Promise<ProviderDetection>;
  /** Models advertised by the installed CLI, when it offers a catalog. */
  listModels?(): Promise<ProviderModelCatalog>;
  /** Read the native title without a model request. Missing titles stay unknown. */
  sessionTitle?(nativeSessionId: string): Promise<string | undefined>;
  /** Whether the CLI is logged in and ready to accept a turn. */
  connectionStatus(): Promise<ProviderConnectionStatus>;
  /** Open one long-lived runtime, creating or resuming a native conversation.
   * Configuration is fixed until closeSession; never silently replace a failed resume.
   */
  startSession(ctx: StartSessionContext): Promise<string>;
  /** Run one turn in the existing runtime. A result ends the turn, not the session.
   * Iterator completion must release turn resources before the next sendTurn.
   */
  sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent>;
  /** Interrupt the running turn. Normal cancellation retains the session. */
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
