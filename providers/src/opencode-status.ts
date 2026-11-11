import { setTimeout as delay } from "node:timers/promises";

export interface OpencodeStatus {
  state: "waiting" | "retrying";
  attempt?: number;
  overloaded?: boolean;
}

/** Older ACP versions omit retry notifications. Read the same process's session status.
 * Never forward provider response bodies/messages or statuses of other sessions.
 */
export async function watchOpencodeStatus(options: {
  port: number; password: string; cwd: string; sessionId: string; signal: AbortSignal;
  onStatus: (status: OpencodeStatus) => void;
}): Promise<void> {
  const { signal } = options;
  let previous = "";
  while (!signal.aborted) {
    try {
      const response = await fetch(`http://127.0.0.1:${options.port}/session/status?directory=${encodeURIComponent(options.cwd)}`, {
        headers: { authorization: `Basic ${Buffer.from(`excel-partner:${options.password}`).toString("base64")}` },
        signal: AbortSignal.any([signal, AbortSignal.timeout(5000)]),
      });
      if (response.status === 404 || response.status === 405) return;
      if (response.ok) {
        const data = await response.json() as Record<string, any>;
        const current = data?.[options.sessionId];
        if (current?.type === "retry" || current?.type === "busy") {
          const status: OpencodeStatus = current.type === "busy" ? { state: "waiting" } : {
            state: "retrying",
            ...(Number.isSafeInteger(current.attempt) && current.attempt >= 0 ? { attempt: current.attempt } : {}),
            overloaded: typeof current.message === "string" && /overload|too many requests|rate.limit|\b429\b/i.test(current.message),
          };
          const key = JSON.stringify(status);
          if (previous !== key && !signal.aborted) { previous = key; options.onStatus(status); }
        }
      }
    } catch { if (signal.aborted) return; /* Optional status polling must not interrupt ACP. */ }
    try { await delay(1000, undefined, { signal }); } catch { if (signal.aborted) return; }
  }
}
