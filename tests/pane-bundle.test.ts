// @vitest-environment jsdom
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { describe, expect, it, vi } from "vitest";

// The pane ships as one esbuild bundle in which the chat modules import each other. Run that bundle, not the sources, so an
// initialisation-order mistake between modules shows up here and not only in the user's Office.
const root = path.join(__dirname, "..");
const g = globalThis as Record<string, any>;

describe("the bundled task pane", () => {
  it("starts and registers with Office", () => {
    const out = path.join(root, "tmp", "pane-bundle-test.js");
    fs.mkdirSync(path.dirname(out), { recursive: true });
    execFileSync(process.execPath, [path.join(root, "node_modules", "esbuild", "bin", "esbuild"), path.join(root, "frontend", "taskpane", "main.ts"), "--bundle", "--format=iife", "--target=es2020", `--outfile=${out}`, "--log-level=error"]);
    document.documentElement.innerHTML = fs.readFileSync(path.join(root, "frontend", "taskpane", "index.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
    let ready: ((info: { host: string }) => void) | undefined;
    g.Office = {
      onReady: (callback: (info: { host: string }) => void) => { ready = callback; return Promise.resolve({ host: "Word" }); },
      HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
      context: { requirements: { isSetSupported: () => true }, host: "Word", platform: "PC", displayLanguage: "en-US", ui: {},
        document: { settings: { get: () => null, set: () => undefined, saveAsync: () => undefined }, addHandlerAsync: vi.fn(), removeHandlerAsync: vi.fn() } },
      EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    };
    g.ResizeObserver = class { observe() {} disconnect() {} unobserve() {} };
    g.matchMedia ??= () => ({ matches: false, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} });
    window.matchMedia ??= g.matchMedia;
    g.WebSocket = class { static OPEN = 1; readyState = 0; send() {} close() {} };
    g.fetch = vi.fn(async () => new Response("{}", { status: 200 }));
    // eslint-disable-next-line no-new-func
    new Function(fs.readFileSync(out, "utf8"))();
    expect(ready, "the bundle must register Office.onReady").toBeTypeOf("function");
    expect(() => ready!({ host: "Excel" })).not.toThrow();
    fs.rmSync(out, { force: true });
  });
});
