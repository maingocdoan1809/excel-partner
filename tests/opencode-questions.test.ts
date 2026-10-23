import http from "node:http";
import { afterEach, describe, expect, it } from "vitest";
import { watchOpencodeQuestions } from "../providers/src/opencode-questions.js";
import type { UserInputResponse } from "../protocol/src/index.js";

let server: http.Server | undefined;
afterEach(async () => { server?.closeAllConnections(); if (server) await new Promise<void>(resolve => server!.close(() => resolve())); });

describe("OpenCode native question API", () => {
  it.each([false, true])("returns native answers or rejects a skipped question (skip=%j)", async skip => {
    const controller = new AbortController();
    const calls: Array<{ url: string; auth: string | undefined; body: any }> = [];
    let finish!: () => void;
    const answered = new Promise<void>(resolve => { finish = resolve; });
    server = http.createServer((req, res) => {
      let raw = "";
      req.on("data", chunk => { raw += chunk; });
      req.on("end", () => {
        calls.push({ url: req.url!, auth: req.headers.authorization, body: raw ? JSON.parse(raw) : null });
        res.setHeader("content-type", "application/json");
        if (req.method === "GET") res.end(JSON.stringify([
          { id: "foreign", sessionID: "another-session", questions: [{ question: "Never display this" }] },
          { id: "question-id", sessionID: "mine", questions: [{ header: "Columns", question: "Which columns?", options: [{ label: "A" }, { label: "B" }], multiple: true }] },
        ]));
        else { res.end("true"); finish(); }
      });
    });
    await new Promise<void>(resolve => server!.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    if (!address || typeof address === "string") throw new Error("No port");
    const running = watchOpencodeQuestions({ port: address.port, password: "test-password", cwd: "/tmp/test-workbook", sessionId: "mine", signal: controller.signal,
      ask: async (request): Promise<UserInputResponse> => {
        expect(request.questions).toHaveLength(1);
        expect(request.questions[0]).toMatchObject({ multiSelect: true, question: "Which columns?" });
        return { cancelled: skip, answers: skip ? {} : { "0": ["A", "B"] } };
      },
    }).catch(error => { if (!controller.signal.aborted) throw error; });
    await answered;
    controller.abort();
    await running;
    expect(calls.every(call => call.auth === `Basic ${Buffer.from("excel-partner:test-password").toString("base64")}`)).toBe(true);
    expect(calls.at(-1)).toMatchObject({ url: `/question/question-id/${skip ? "reject" : "reply"}?directory=%2Ftmp%2Ftest-workbook`, body: skip ? {} : { answers: [["A", "B"]] } });
  });
});
