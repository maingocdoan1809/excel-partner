import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderDetection,
  ProviderEvent,
  ProviderId,
  PromptSelectionSnapshot,
} from "../../protocol/src/index.js";
import type { ProviderOption } from "../../protocol/src/provider-options.js";
import type { LoginAccount, LoginChoice, ProviderLoginStatus } from "../../protocol/src/provider-login.js";

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
   * (EXCEL_BRIDGE_URL, EXCEL_SESSION_ID, EXCEL_SESSION_TOKEN, OFFICE_HOST and, for the
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
  selection: PromptSelectionSnapshot | null;
  /** Aborted when the user hits stop or the pane disconnects. */
  signal: AbortSignal;
}

/** One sign-in attempt. The bridge shows its status and hands it the code the sign-in page returned. */
export interface ProviderLoginAttempt {
  status(): ProviderLoginStatus;
  /** Hands over the code the sign-in page showed; throws when no code is expected. */
  submitCode(code: string): void;
  cancel(): void;
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
  runtimeRevision?(): string;
  listCommands?(providerSessionId?: string): Promise<import("../../protocol/src/commands.js").ProviderCommand[]>;
  readonly id: ProviderId;
  readonly capabilities: ProviderCapabilities;
  /** Extra instructions for how this provider reaches the client's MCP tools, when its way differs from the usual one. */
  readonly clientToolNotes?: string;
  /** Whether the CLI is installed and what it supports. */
  detect(force?: boolean): Promise<ProviderDetection>;
  /**
   * Proves a turn can start: the CLI is signed in and a session opens with the given MCP tools registered.
   * Sends no prompt. Rejects with a ProviderLoginError when the person still has to sign in.
   */
  verifyReady(mcp: McpEndpoint): Promise<void>;
  /** Starts a sign-in with the provider's own CLI. `choice` is only for providers that list accounts (login kind `account`). */
  startLogin(choice?: LoginChoice): ProviderLoginAttempt;
  /** Installs the CLI the release names for this provider and makes the provider use it. */
  install(): Promise<void>;
  /** Stops using the CLI the product installed for this provider. */
  uninstall(): Promise<void>;
  /** Signs the CLI out (`accountId` names which account, for providers that list them). Rejects when the provider has no sign-out. */
  logout(accountId?: string): Promise<void>;
  /** The accounts to choose from; only providers whose login kind is `account`. */
  loginAccounts?(): Promise<LoginAccount[]>;
  /** Settings the CLI offers for this session beyond model and effort (a mode, for example), as it last reported them. */
  sessionOptions?(providerSessionId: string): ProviderOption[];
  /** Changes one of them in the running session and returns the options as the CLI now reports them. */
  setSessionOption?(providerSessionId: string, optionId: string, value: string): Promise<ProviderOption[]>;
  /** Models advertised by the installed CLI, when it offers a catalog. */
  listModels?(): Promise<ProviderModelCatalog>;
  /** Read the native title without a model request. Missing titles stay unknown. */
  sessionTitle?(nativeSessionId: string, options?: { cheapOnly?: boolean }): Promise<string | undefined>;
  /**
   * What the CLI itself recorded for a conversation: the session's total token usage. Only for a CLI whose per-turn
   * figure is not complete or costs a process to read; it is asked for when someone looks at the usage, not after every turn.
   */
  sessionUsage?(nativeSessionId: string): Promise<{ usage: import("../../protocol/src/index.js").TokenUsage; title?: string }>;
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
  /**
   * Hand a message from the person to the turn that is running now, the way the CLI itself takes one at its next step.
   * Resolves true once the CLI has it and will answer it inside the running turn (the turn's own events carry the
   * answer); false when there is no turn to join or the CLI cannot take it, and the caller queues it instead.
   * Optional: an adapter without it always gets the queue.
   */
  steerTurn?(providerSessionId: string, text: string): Promise<boolean>;
  /** End the session and release the CLI process. */
  closeSession(providerSessionId: string): Promise<void>;
  /** Release all resources; called on bridge shutdown. */
  dispose(): Promise<void>;
}

export interface ProviderRuntime {
  /** Immutable per-profile child environment; credentials remain provider-owned. */
  env?: NodeJS.ProcessEnv;
  bridgeUrl: string;
  mcp: { command: string; args: string[] };
  logger: { debug(msg: string, extra?: unknown): void; warn(msg: string, extra?: unknown): void };
}
