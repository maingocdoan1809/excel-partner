import { describe, expect, it, vi } from "vitest";
import { acpPermissionOutcome, contentText, toolNameOf } from "../providers/src/acp.js";
import { effortConfig } from "../providers/src/effort.js";
import { OpencodeAdapter, isOpencodeV2, parseOpencodeModels } from "../providers/src/opencode.js";

const PERMISSION_OPTIONS = [
  { optionId: "once", kind: "allow_once", name: "Allow once" },
  { optionId: "always", kind: "allow_always", name: "Always allow" },
  { optionId: "reject", kind: "reject_once", name: "Reject" },
];

describe("OpenCode model catalog", () => {
  it("reads the model config option and ignores the other options", () => {
    const catalog = parseOpencodeModels([
      { id: "model", category: "model", type: "select", currentValue: "opencode/big-pickle", options: [
        { value: "anthropic/claude-sonnet-5-5", name: "Anthropic/Claude Sonnet 5.5" },
        { value: "opencode/big-pickle", name: "OpenCode/Big Pickle" },
      ] },
      { id: "mode", category: "mode", type: "select", currentValue: "build", options: [{ value: "build", name: "Build" }] },
    ]);
    expect(catalog.models).toEqual([
      { id: "anthropic/claude-sonnet-5-5", name: "Anthropic/Claude Sonnet 5.5" },
      { id: "opencode/big-pickle", name: "OpenCode/Big Pickle" },
    ]);
    expect(catalog.defaultModelId).toBe("opencode/big-pickle");
  });

  it("falls back to the model id when the CLI reports no display name", () => {
    expect(parseOpencodeModels([{ id: "model", options: [{ value: "opencode/beam-free" }] }]).models)
      .toEqual([{ id: "opencode/beam-free", name: "opencode/beam-free" }]);
  });

  it("returns an empty catalog when the CLI offers no models", () => {
    expect(parseOpencodeModels(undefined)).toEqual({ models: [], defaultModelId: null });
  });
});

describe("OpenCode effort option", () => {
  it("maps the requested effort onto the model variants", () => {
    const options = [
      { id: "effort", category: "thought_level", type: "select", currentValue: "medium", options: [
        { value: "low" }, { value: "medium" }, { value: "high" },
      ] },
    ];
    expect(effortConfig(options, "high")).toEqual({ configId: "effort", value: "high" });
  });

  it("leaves the effort alone for models without variants", () => {
    expect(effortConfig([{ id: "model", category: "model", type: "select", options: [{ value: "opencode/big-pickle" }] }], "high")).toBeNull();
  });
});

describe("ACP tool names and permissions", () => {
  it("recognises the Excel MCP tools of every CLI naming style", () => {
    expect(toolNameOf({ title: "excel_get_selection" })).toBe("get_selection");
    expect(toolNameOf({ title: "excel-get_selection" })).toBe("get_selection");
    expect(toolNameOf({ title: "excel/read_range" }, true)).toEqual({ name: "read_range", isExcel: true });
    expect(toolNameOf({ title: "echo hi" }, true).isExcel).toBe(false);
  });

  it("allows Excel tools and rejects everything else", () => {
    expect(acpPermissionOutcome({ toolCall: { title: "excel_write_range" }, options: PERMISSION_OPTIONS }))
      .toEqual({ outcome: { outcome: "selected", optionId: "once" } });
    expect(acpPermissionOutcome({ toolCall: { title: "echo hi" }, options: PERMISSION_OPTIONS }))
      .toEqual({ outcome: { outcome: "selected", optionId: "reject" } });
  });

  it("cancels when the CLI offers no matching option", () => {
    expect(acpPermissionOutcome({ toolCall: { title: "echo hi" }, options: [] }))
      .toEqual({ outcome: { outcome: "cancelled" } });
  });

  it("flattens the content of a finished tool call", () => {
    expect(contentText({ content: [{ content: { type: "text", text: "A1:B2" } }] })).toBe("A1:B2");
    expect(contentText({ rawOutput: "done" })).toBe("done");
  });
});


describe("OpenCode pane permission routing", () => {
  const adapter = new OpencodeAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  const request = { toolCall: { title: "bash", rawInput: { command: "echo hello" } }, options: PERMISSION_OPTIONS };
  const ask = (params: unknown, authorize?: (args: unknown) => Promise<boolean>) =>
    (adapter as any).onServerRequest("session/request_permission", params, authorize);
  it("uses the pane approval for native tools", async () => {
    let received: unknown;
    expect(await ask(request, async args => { received = args; return true; }))
      .toEqual({ outcome: { outcome: "selected", optionId: "once" } });
    expect(received).toEqual(request.toolCall);
    expect(await ask(request, async () => false)).toEqual({ outcome: { outcome: "selected", optionId: "reject" } });
  });
  it("rejects missing or failed approval and leaves Excel approval in its own gateway", async () => {
    expect(await ask(request)).toEqual({ outcome: { outcome: "selected", optionId: "reject" } });
    expect(await ask(request, async () => { throw new Error("disconnected"); })).toEqual({ outcome: { outcome: "selected", optionId: "reject" } });
    expect(await ask({ ...request, toolCall: { title: "excel_write_range" } }, async () => { throw new Error("should not ask"); }))
      .toEqual({ outcome: { outcome: "selected", optionId: "once" } });
  });
});


describe("OpenCode ACP retry diagnostics", () => {
  it("reports retry status without exposing the upstream body or message", () => {
    const adapter = new OpencodeAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    const diagnostic = vi.fn();
    const session = { acpSessionId: "native", active: {}, diagnostic };
    const notify = (adapter as any).onNotification.bind(adapter);
    const update = { sessionUpdate: "session_info_update", _meta: { "opencode/retry": {
      attempt: 1, nextRetryAt: "2026-10-24T08:00:00Z",
      error: { type: "provider.rate-limit", status: 429, message: "PRIVATE_MESSAGE", response: { body: "PRIVATE_BODY" } },
    } } };
    notify(session, "session/update", { sessionId: "different", update });
    expect(diagnostic).not.toHaveBeenCalled();
    notify(session, "session/update", { sessionId: "native", update });
    expect(diagnostic).toHaveBeenLastCalledWith({ phase: "model_retry_scheduled", attempt: 1, status: 429,
      errorType: "provider.rate-limit", nextRetryAt: "2026-10-24T08:00:00.000Z" });
    notify(session, "session/update", { sessionId: "native", update: { sessionUpdate: "session_info_update", _meta: { "opencode/retry": null } } });
    expect(diagnostic).toHaveBeenLastCalledWith({ phase: "model_retry_started" });
    expect(JSON.stringify(diagnostic.mock.calls)).not.toContain("PRIVATE");
  });
});
