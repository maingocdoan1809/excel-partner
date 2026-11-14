// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { acpOptions, type AcpConfigOption } from "../providers/src/options.js";
import { OpencodeAdapter } from "../providers/src/opencode/adapter.js";
import { setupProviderOptions } from "../frontend/taskpane/ui/provider-options.js";
import type { ProviderOption } from "../protocol/src/provider-options.js";

const model: AcpConfigOption = { id: "model", type: "select", category: "model", currentValue: "a", options: [{ value: "a" }, { value: "b" }] };
const mode: AcpConfigOption = { id: "mode", name: "Mode", type: "select", category: "mode", currentValue: "build", options: [{ value: "build", name: "Build" }, { value: "plan", name: "Plan" }] };
const grouped: AcpConfigOption = { id: "effort", type: "select", category: "thought_level", currentValue: "low", options: [{ options: [{ value: "low" }, { value: "high" }] }] };
const permissions: AcpConfigOption = { id: "allow_all", type: "boolean", category: "permissions", currentValue: "off" };

describe("what a provider passes on of the CLI's own options", () => {
  it("keeps only the categories the provider names, with the CLI's own labels", () => {
    expect(acpOptions([model, mode, grouped, permissions], ["mode"])).toEqual([
      { id: "mode", label: "Mode", value: "build", choices: [{ value: "build", label: "Build" }, { value: "plan", label: "Plan" }] },
    ]);
    expect(acpOptions([model, mode, grouped], ["mode", "thought_level"]).map(option => option.id)).toEqual(["mode", "effort"]);
    expect(acpOptions([grouped], ["thought_level"])[0]!.choices.map(choice => choice.value)).toEqual(["low", "high"]);
  });

  it("skips options with nothing to choose and anything that is not a select", () => {
    expect(acpOptions([{ ...mode, options: [] }, permissions], ["mode", "permissions"])).toEqual([]);
    expect(acpOptions(undefined, ["mode"])).toEqual([]);
  });
});

describe("OpenCode's mode", () => {
  function adapterWith(configOptions: AcpConfigOption[]) {
    const adapter = new OpencodeAdapter({ env: {}, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } } as never);
    const request = vi.fn(async (_method: string, params: { configId: string; value: string }) => ({
      configOptions: configOptions.map(option => option.id === params.configId ? { ...option, currentValue: params.value } : option) }));
    (adapter as unknown as { sessions: Map<string, unknown> }).sessions.set("s1", { configOptions, acpSessionId: "acp", peer: { request } });
    return { adapter, request };
  }

  it("is offered with the model and effort left to their own controls", () => {
    const { adapter } = adapterWith([model, mode]);
    expect(adapter.sessionOptions("s1").map(option => option.id)).toEqual(["mode"]);
    expect(adapter.sessionOptions("missing")).toEqual([]);
  });

  it("is changed in the running session through the CLI, and reported as the CLI now has it", async () => {
    const { adapter, request } = adapterWith([model, mode]);
    const options = await adapter.setSessionOption("s1", "mode", "plan");
    expect(request).toHaveBeenCalledWith("session/set_config_option", { sessionId: "acp", configId: "mode", value: "plan" }, 10_000);
    expect(options[0]).toMatchObject({ id: "mode", value: "plan" });
  });

  it("refuses a setting or a value the CLI does not offer, without asking the CLI", async () => {
    const { adapter, request } = adapterWith([model, mode]);
    await expect(adapter.setSessionOption("s1", "model", "b")).rejects.toThrow(/does not offer/);
    await expect(adapter.setSessionOption("s1", "mode", "turbo")).rejects.toThrow(/does not offer/);
    await expect(adapter.setSessionOption("nope", "mode", "plan")).rejects.toThrow(/not running/);
    expect(request).not.toHaveBeenCalled();
  });
});

describe("the pane's control for the provider's options", () => {
  const options: ProviderOption[] = [{ id: "mode", label: "Mode", value: "build", choices: [{ value: "build", label: "Build" }, { value: "plan", label: "Plan" }] }];
  function control(overrides: { busy?: boolean; change?: (id: string, value: string) => Promise<ProviderOption[]> } = {}) {
    const container = document.createElement("div");
    const fail = vi.fn();
    const change = overrides.change ?? vi.fn(async (_id: string, value: string) => [{ ...options[0]!, value }]);
    const view = setupProviderOptions({ container, change, isBusy: () => overrides.busy ?? false, fail });
    return { container, fail, change, view };
  }
  const rows = (container: HTMLElement) => [...container.querySelectorAll<HTMLButtonElement>(".provider-choice")];
  const checked = (container: HTMLElement) => rows(container).find(row => row.getAttribute("aria-checked") === "true")?.dataset.value;

  it("is hidden until the provider reports options, and hidden again when they no longer apply", () => {
    const { container, view } = control();
    expect(container.hidden).toBe(true);
    view.render(options);
    expect(container.hidden).toBe(false);
    expect(container.querySelector(".provider-option-title")?.textContent).toBe("Mode");
    expect(rows(container).map(row => row.textContent)).toEqual(["Build", "Plan"]);
    expect(checked(container)).toBe("build");
    view.clear();
    expect(container.hidden).toBe(true);
  });

  it("sends the choice back to the provider and shows what the provider answers", async () => {
    const { container, view, change } = control();
    view.render(options);
    rows(container)[1]!.click();
    await vi.waitFor(() => expect(change).toHaveBeenCalledWith("mode", "plan"));
    await vi.waitFor(() => expect(checked(container)).toBe("plan"));
  });

  it("does nothing when the choice already in force is picked", () => {
    const { container, view, change } = control();
    view.render(options);
    rows(container)[0]!.click();
    expect(change).not.toHaveBeenCalled();
  });

  it("keeps the old choice and says why when the provider refuses", async () => {
    const { container, view, fail } = control({ change: async () => { throw new Error("Wait for the current response to finish."); } });
    view.render(options);
    rows(container)[1]!.click();
    await vi.waitFor(() => expect(fail).toHaveBeenCalledWith("Wait for the current response to finish."));
    expect(checked(container)).toBe("build");
  });

  it("cannot be used while a reply is running", () => {
    const { container, view } = control({ busy: true });
    view.render(options);
    expect(rows(container).every(row => row.disabled)).toBe(true);
  });
});
