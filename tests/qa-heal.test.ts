import { describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { healingHint } from "../bridge/src/qa-tools.js";
import { loadConfig, saveConfig } from "../qa/src/config.js";
import { HARNESS_FILES, templatesDir } from "../qa/src/harness.js";
import { approve, saveScript } from "../qa/src/manifest.js";
import { GOOD_TEST, makeProject } from "./qa-helpers.js";

const state = (extra: Record<string, unknown> = {}) => ({ url: "/web", busy: false, dialogs: [], invalid: [], notices: [], ...extra });
const failed = (pageState?: unknown, failedStep = "locator.click") => ({ status: "failed" as const, pageState: pageState as never, failedStep });

describe("first reading of a failed script", () => {
  it("points at an extra dialog and names its buttons", () => {
    const hint = healingHint(failed(state({ dialogs: [{ title: "Confirmation", text: "Are you sure?", buttons: ["Ok", "Cancel"] }] })));
    expect(hint).toContain('"Confirmation"');
    expect(hint).toContain("Ok, Cancel");
    expect(hint).toContain("extra step");
  });

  it("tells a loading page from a real application message", () => {
    expect(healingHint(failed(state({ busy: true })))).toContain("still loading");
    for (const extra of [{ invalid: ["partner_id"] }, { notices: ["Missing required field"] }]) {
      const hint = healingHint(failed(state(extra)));
      expect(hint).toContain("real difference in behaviour");
      expect(hint).toContain("instead of changing the script");
    }
  });

  it("falls back to the locator and never suggests loosening an assertion", () => {
    const hint = healingHint(failed(state()));
    expect(hint).toContain('after "locator.click" failed');
    expect(hint).toContain("Never loosen an assertion");
  });

  it("copes with no recorded state and with tests that did not fail", () => {
    expect(healingHint(failed(undefined))).toContain("No page state was recorded");
    expect(healingHint({ status: "passed", pageState: undefined, failedStep: undefined })).toBe("This test did not fail.");
  });
});

describe("what the harness records when a script fails", () => {
  const harness = (file: string) => fs.readFileSync(path.join(templatesDir(), "project", "automation", file), "utf8");

  it("ships the page-state helper as a hash-checked harness file and loads it in every test", () => {
    expect(HARNESS_FILES).toContain("_harness/page-state.js");
    expect(harness("_harness/fixtures.js")).toContain("import { collectPageState } from './page-state.js'");
    expect(harness("_harness/fixtures.js")).toContain("{ auto: true }");
    expect(harness("_harness/fixtures.js")).toContain("testInfo.attach('page-state'");
  });

  it("reads the helper's output into the results, redacted, with the failing line and step", () => {
    const reporter = harness("_harness/reporter.js");
    for (const needle of ["page-state", "redactDeep", "failedAt", "failedStep", "snippet", "result.error.location"]) expect(reporter, needle).toContain(needle);
  });

  it("does not use anything the validator forbids, because scripts import only fixtures, not this file", () => {
    const helper = harness("_harness/page-state.js");
    expect(helper).not.toMatch(/\bfetch\(|child_process|require\(|process\./);
    expect(helper).toContain("Promise.race"); // a hung page must not stall the report
  });
});

describe("running a script the assistant has just repaired", () => {
  function fixture(healWithoutPrompt: boolean) {
    const { dir, config } = makeProject();
    saveConfig(dir, { ...config, run: { ...config.run, healWithoutPrompt } });
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("heal", null);
    Object.defineProperty(session, "isConnected", { get: () => true });
    const asked: Array<{ tool: string; args: Record<string, unknown> }> = [];
    vi.spyOn(session, "requestTool").mockImplementation(async (tool, args) => { asked.push({ tool, args: args as Record<string, unknown> }); return { ok: true, data: { approved: true } }; });
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, null, null, null, undefined, () => dir);
    const cfg = loadConfig(dir);
    const repaired = () => {
      saveScript(dir, cfg, { tcId: "TC_001", title: "x", code: GOOD_TEST, role: "admin" });
      approve(dir, cfg, "TC_001", "person");
      saveScript(dir, cfg, { tcId: "TC_001", title: "x", code: GOOD_TEST.replace("/web'", "/web#fixed'"), role: "admin" }); // now "modified"
    };
    const run = (tcIds: string[]) => gateway.invoke({ sessionId: session.id, tool: "qa_run", args: { tcIds, draft: true } });
    return { asked, repaired, run };
  }

  it("always asks by default, even for a script that had been approved", async () => {
    const { asked, repaired, run } = fixture(false);
    repaired();
    await run(["TC_001"]);
    expect(asked.at(-1)).toMatchObject({ tool: "__authorize_qa", args: { action: "run_draft", forceAsk: true } });
  });

  it("skips the question for repaired scripts only when the project owner allowed it", async () => {
    const { asked, repaired, run } = fixture(true);
    repaired();
    await run(["TC_001"]);
    expect(asked.at(-1)!.args).toMatchObject({ action: "run_draft", forceAsk: false });
  });

  it("still asks for a script that was never approved, or when one of several is brand new", async () => {
    const { asked, repaired, run } = fixture(true);
    await run(["TC_000"]); // the example script is a draft that nobody approved
    expect(asked.at(-1)!.args.forceAsk).toBe(true);
    repaired();
    await run(["TC_001", "TC_000"]);
    expect(asked.at(-1)!.args.forceAsk).toBe(true);
  });
});
