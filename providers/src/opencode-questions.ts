import net from "node:net";
import { setTimeout as delay } from "node:timers/promises";
import { UserInputRequestSchema } from "../../protocol/src/index.js";
import { askProviderQuestions, type RequestUserInput } from "./user-input.js";

export async function reserveOpencodePort(): Promise<number> {
  const server = net.createServer();
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      const port = typeof address === "object" && address ? address.port : 0;
      server.close(error => error ? reject(error) : resolve(port));
    });
  });
}

/** ACP does not forward question events; use the same process's native question API. */
export async function watchOpencodeQuestions(options: {
  port: number; password: string; cwd: string; sessionId: string; ask: RequestUserInput; signal: AbortSignal;
}): Promise<void> {
  const { signal } = options;
  const query = `?directory=${encodeURIComponent(options.cwd)}`;
  const headers = { authorization: `Basic ${Buffer.from(`excel-partner:${options.password}`).toString("base64")}`, "content-type": "application/json" };
  const call = async (route: string, body?: unknown) => {
    const response = await fetch(`http://127.0.0.1:${options.port}${route}${query}`, {
      headers, method: body === undefined ? "GET" : "POST", body: body === undefined ? undefined : JSON.stringify(body),
      signal: AbortSignal.any([signal, AbortSignal.timeout(10_000)]),
    });
    if (!response.ok) throw new Error(`OpenCode question API returned HTTP ${response.status}`);
    return response.json();
  };
  const handled = new Set<string>();
  while (!signal.aborted) {
    const pending = await call("/question") as Array<{ id: string; sessionID: string; questions: Array<Record<string, any>> }>;
    for (const item of pending) {
      if (signal.aborted) return;
      if (item.sessionID !== options.sessionId || handled.has(item.id)) continue;
      handled.add(item.id);
      const request = UserInputRequestSchema.parse({ questions: item.questions.map((q, index) => ({
        ...q, id: String(index), multiSelect: q.multiple ?? false, allowFreeform: q.custom !== false,
      })) });
      const response = await askProviderQuestions(request, options.ask, signal);
      if (signal.aborted) return;
      // A question may have been withdrawn while the user was answering.
      const current = await call("/question") as Array<{ id: string; sessionID: string }>;
      if (!current.some(q => q.id === item.id && q.sessionID === options.sessionId)) continue;
      await call(`/question/${encodeURIComponent(item.id)}/${response.cancelled ? "reject" : "reply"}`,
        response.cancelled ? {} : { answers: request.questions.map(q => response.answers[q.id]) });
    }
    await delay(500, undefined, { signal });
  }
}
