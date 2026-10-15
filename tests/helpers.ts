import net from "node:net";
import WebSocket from "ws";
import {
  decodeServerMessage,
  encodeMessage,
  type ClientMessage,
  type ServerMessage,
} from "../protocol/src/index.js";

export function getFreePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      const port = typeof address === "object" && address ? address.port : 0;
      server.close(() => resolve(port));
    });
  });
}

export class PaneClient {
  readonly received: ServerMessage[] = [];
  /** Optional observer invoked for every incoming message. */
  onMessage: ((message: ServerMessage) => void) | null = null;
  private readonly waiters: Array<{
    predicate: (message: ServerMessage, index: number) => boolean;
    resolve: (message: ServerMessage) => void;
  }> = [];

  private constructor(private readonly ws: WebSocket) {
    ws.on("message", (data) => {
      const message = decodeServerMessage(data.toString());
      this.received.push(message);
      this.onMessage?.(message);
      for (let i = this.waiters.length - 1; i >= 0; i--) {
        const waiter = this.waiters[i]!;
        if (waiter.predicate(message, this.received.length - 1)) {
          this.waiters.splice(i, 1);
          waiter.resolve(message);
        }
      }
    });
  }

  static connect(url: string, options: { ca?: string | Buffer } = {}): Promise<PaneClient> {
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(url, options.ca ? { ca: options.ca } : undefined);
      const client = new PaneClient(ws);
      ws.once("open", () => resolve(client));
      ws.once("error", reject);
    });
  }

  send(message: ClientMessage): void {
    this.ws.send(encodeMessage(message));
  }

  /** Resolve with the first message matching `predicate` (past or future). */
  waitFor(predicate: (message: ServerMessage, index: number) => boolean, timeoutMs = 5000): Promise<ServerMessage> {
    const existing = this.received.find(predicate);
    if (existing) return Promise.resolve(existing);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`timed out waiting for a matching message`));
      }, timeoutMs);
      this.waiters.push({
        predicate,
        resolve: (message) => {
          clearTimeout(timer);
          resolve(message);
        },
      });
    });
  }

  close(): Promise<void> {
    return new Promise((resolve) => {
      if (this.ws.readyState === WebSocket.CLOSED) return resolve();
      this.ws.once("close", () => resolve());
      this.ws.close();
    });
  }
}
