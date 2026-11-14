import { newSession, type StoredMessage, type StoredSession, type StoredTool } from "../ui/sessions.js";
import { BridgeClient, type UploadedAttachment } from "../bridge-client.js";
import type { ProviderId } from "../../../protocol/src/index.js";
import type { PermissionMode } from "../permissions.js";
import { setupComposerMenu } from "../ui/composer-menu.js";
import { setupSelectionChip } from "../ui/selection-chip.js";
import { loadProvider } from "../ui/provider-picker.js";
import { boot } from "./boot.js";
import { store } from "./services.js";

export const PERMISSION_MODE_KEY = "excelPartner.permissionMode";
export interface PendingPrompt {
  id: string;
  parentRequestId: string;
  prompt: string;
  attachments: typeof pendingAttachments;
  message: StoredMessage;
  state: "pending" | "next_turn" | "held";
  /** Why the bridge refused it, shown beside "Not sent". */
  reason?: string;
}

export const pendingPrompts = new Map<string, PendingPrompt>();
export const liveTools = new Map<string, { card: HTMLElement; tool: StoredTool }>();
export const pendingAttachments: Array<{ reference: UploadedAttachment; name: string }> = [];

/**
 * Where the chat is right now: the open chat, the running request and what is on screen for it.
 * One object, so every module reads and changes the same values.
 */
export interface AppState {
  /** Opened when Office is ready (chat/boot.ts). */
  client: BridgeClient;
  provider: ProviderId;
  activeModel: string | null | undefined;
  activeEffort: string | undefined;
  permissionMode: PermissionMode;
  session: StoredSession;
  /** After a reload, the chat to bring back once the history store is connected. */
  restoreChatId: string | null;
  /** Where the running request started in the chat, so a reply is bound to the right native session. */
  requestMessageIndex: number;
  requestChangeIndex: number;
  requestStartedAt: number;
  /** The provider process starts empty after a reload or when a past chat is continued; replay the transcript once. */
  needsContext: boolean;
  currentRequestId: string | null;
  requestSession: StoredSession | null;
  cancelRequested: boolean;
  promptDispatched: boolean;
  autoQueueAllowed: boolean;
  resolvingCommand: boolean;
  /** The assistant message being streamed, and the open tool disclosure of the turn. */
  live: { element: HTMLElement; message: StoredMessage } | null;
  liveToolGroup: { element: HTMLDetailsElement; message: StoredMessage } | null;
  /** The other workbooks open with this add-in (not this one), as the bridge last listed them. */
  otherFiles: string[];
  /** Set once the assistant has used the connected-file tools in the running request: what it says may be about another file. */
  turnTouchedOtherFiles: boolean;
  /** The names of this workbook's sheets, as last read from Excel (so a reply that names one is not guessed at). */
  sheetNames: string[];
  permissionPicker: ReturnType<typeof setupComposerMenu> | undefined;
  selectionChip: ReturnType<typeof setupSelectionChip> | undefined;
  selectedSlashToken: string | null;
}

const initialProvider = loadProvider();
export const app: AppState = {
  client: undefined as unknown as BridgeClient,
  provider: initialProvider,
  activeModel: undefined,
  activeEffort: undefined,
  permissionMode: localStorage.getItem(PERMISSION_MODE_KEY) === "auto" ? "auto" : "ask",
  session: newSession(initialProvider),
  restoreChatId: null,
  requestMessageIndex: 0,
  requestChangeIndex: 0,
  requestStartedAt: 0,
  needsContext: false,
  currentRequestId: null,
  requestSession: null,
  cancelRequested: false,
  promptDispatched: false,
  autoQueueAllowed: true,
  resolvingCommand: false,
  live: null,
  liveToolGroup: null,
  otherFiles: [],
  turnTouchedOtherFiles: false,
  sheetNames: [],
  permissionPicker: undefined,
  selectionChip: undefined,
  selectedSlashToken: null,
};
