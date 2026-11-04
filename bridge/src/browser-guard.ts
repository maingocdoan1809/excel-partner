import { FRAME_AFTER } from "./browser-frames.js";

/**
 * Stops the assistant from repeating a browser action that keeps failing. A model that cannot see why a click
 * failed will often try the very same click again and again; after a couple of identical failures the tool
 * refuses and says what to do instead. Reading tools (snapshot, screenshot, console) are never blocked, so the
 * assistant can always look around.
 */
export const MAX_SAME_FAILURES = 2;
export const MAX_CONSECUTIVE_FAILURES = 5;

/** Actions that change the page; the same set that gets a live picture. */
export const GUARDED_TOOLS = FRAME_AFTER;

/** Free-text descriptions change between retries; the thing being acted on does not. */
const IGNORED_ARGS = new Set(["element", "description", "intent", "timeout"]);

export function actionSignature(tool: string, args: Record<string, unknown> = {}): string {
  const kept = Object.entries(args).filter(([key]) => !IGNORED_ARGS.has(key)).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return `${tool}:${JSON.stringify(kept)}`;
}

export interface GuardVerdict { reason: "same" | "streak"; attempts: number }
interface GuardState { turn?: string; signature?: string; same: number; consecutive: number }

export class BrowserGuard {
  private readonly states = new Map<string, GuardState>();

  private state(sessionId: string, turn?: string): GuardState {
    let state = this.states.get(sessionId);
    // A new request (the person answered, or asked again) gets a fresh start.
    if (!state || (turn !== undefined && state.turn !== turn)) { state = { turn, same: 0, consecutive: 0 }; this.states.set(sessionId, state); }
    return state;
  }

  /** Null = go ahead. Otherwise the call must not run. Refusals are not counted, so they keep being refused. */
  check(sessionId: string, tool: string, args: Record<string, unknown>, turn?: string): GuardVerdict | null {
    if (!GUARDED_TOOLS.has(tool)) return null;
    const state = this.state(sessionId, turn);
    if (state.consecutive >= MAX_CONSECUTIVE_FAILURES) return { reason: "streak", attempts: state.consecutive };
    if (state.signature === actionSignature(tool, args) && state.same >= MAX_SAME_FAILURES) return { reason: "same", attempts: state.same };
    return null;
  }

  /** Returns how many times in a row this exact action has now failed (0 after a success). */
  record(sessionId: string, tool: string, args: Record<string, unknown>, ok: boolean, turn?: string): { same: number; consecutive: number } {
    if (!GUARDED_TOOLS.has(tool)) return { same: 0, consecutive: this.states.get(sessionId)?.consecutive ?? 0 };
    const state = this.state(sessionId, turn);
    if (ok) { state.signature = undefined; state.same = 0; state.consecutive = 0; return { same: 0, consecutive: 0 }; }
    const signature = actionSignature(tool, args);
    if (state.signature === signature) state.same++;
    else { state.signature = signature; state.same = 1; }
    state.consecutive++;
    return { same: state.same, consecutive: state.consecutive };
  }

  clear(sessionId: string): void { this.states.delete(sessionId); }
}
