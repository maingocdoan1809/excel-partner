// @vitest-environment jsdom
import fs from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setupUpdatePanel, type UpdateClient, type UpdateInfo } from "../addin/taskpane/ui/update-panel.js";

const info = (extra: Partial<UpdateInfo> = {}): UpdateInfo => ({
  current: "0.1.19", latest: "0.1.20", updateAvailable: true, canUpdate: true, status: { state: "idle" }, logFile: "C:\\data\\update.log", ...extra,
});

const el = (id: string) => document.getElementById(id) as HTMLElement;
const button = (id: string) => el(id) as HTMLButtonElement;
const flush = async () => { for (let i = 0; i < 6; i++) await Promise.resolve(); };

function storage() {
  const data = new Map<string, string>();
  return { getItem: (k: string) => data.get(k) ?? null, setItem: (k: string, v: string) => void data.set(k, v), removeItem: (k: string) => void data.delete(k), data };
}

function setup(replies: Array<UpdateInfo | Error> = [info()], over: Partial<UpdateClient> = {}) {
  document.body.innerHTML = fs.readFileSync("addin/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
  const queue = [...replies];
  const client = {
    updateInfo: vi.fn(async () => {
      const next = queue.length > 1 ? queue.shift()! : queue[0]!;
      if (next instanceof Error) throw next;
      return next;
    }),
    startUpdate: vi.fn(async () => undefined),
    ...over,
  };
  const store = storage();
  const reload = vi.fn();
  const announce = vi.fn();
  const panel = setupUpdatePanel({ getClient: () => client, reload, announce, storage: store, pollMs: 100, waitMs: 1000, confirmMs: 500 });
  return { client, store, reload, announce, panel, queue };
}

beforeEach(() => vi.useFakeTimers());
afterEach(() => vi.useRealTimers());

describe("Settings > Updates", () => {
  it("shows the installed version and says when it is the latest", async () => {
    const { panel } = setup([info({ latest: "0.1.19", updateAvailable: false })]);
    await panel.refresh();
    expect(el("update-version").textContent).toBe("Version 0.1.19");
    expect(el("update-status").textContent).toBe("You have the latest version.");
    expect(el("update-run").hidden).toBe(true);
  });

  it("offers the update, and says Excel can stay open", async () => {
    const { panel } = setup();
    await panel.refresh();
    expect(el("update-run").hidden).toBe(false);
    expect(el("update-run").textContent).toBe("Update to 0.1.20");
    expect(el("update-status").textContent).toContain("Version 0.1.20 is available");
    expect(el("update-status").textContent).toContain("Excel stays open");
  });

  it("explains why an update cannot be done from here instead of offering a button", async () => {
    const { panel } = setup([info({ canUpdate: false, reason: "This copy was not installed with npm install -g." })]);
    await panel.refresh();
    expect(el("update-run").hidden).toBe(true);
    expect(el("update-status").textContent).toContain("Version 0.1.20 is available");
    expect(el("update-status").textContent).toContain("not installed with npm install -g");
  });

  it("says plainly when the check itself failed", async () => {
    const { panel } = setup([info({ latest: null, updateAvailable: false, checkError: "getaddrinfo ENOTFOUND" })]);
    await panel.refresh();
    expect(el("update-status").textContent).toBe("Could not check for updates: getaddrinfo ENOTFOUND");
    expect(el("update-status").dataset.kind).toBe("error");
    const failed = setup([new Error("Bridge is not connected.")]);
    await failed.panel.refresh();
    expect(el("update-status").textContent).toContain("Bridge is not connected.");
  });

  it("reports an earlier update that did not finish, with the log to read", async () => {
    const { panel } = setup([info({ updateAvailable: false, status: { state: "failed", error: "EBUSY: resource busy" } })]);
    await panel.refresh();
    expect(el("update-status").textContent).toContain("did not finish: EBUSY: resource busy");
    expect(el("update-status").textContent).toContain("C:\\data\\update.log");
  });

  it("the Check button asks the registry again", async () => {
    const { panel, client } = setup();
    await panel.refresh();
    button("update-check").click();
    await flush();
    expect(client.updateInfo).toHaveBeenLastCalledWith(true);
  });

  it("needs a second click to start, and forgets the first one after a moment", async () => {
    const { panel, client } = setup();
    await panel.refresh();
    button("update-run").click();
    expect(el("update-run").textContent).toBe("Click again to update");
    expect(client.startUpdate).not.toHaveBeenCalled();
    vi.advanceTimersByTime(600);
    expect(el("update-run").textContent).toBe("Update to 0.1.20");
    button("update-run").click();
    vi.advanceTimersByTime(600);
    button("update-run").click();
    await flush();
    expect(client.startUpdate).not.toHaveBeenCalled();
  });

  it("starts on the second click and follows the update until it is done, then reloads the pane", async () => {
    const running = (message: string) => info({ status: { state: "running", message, from: "0.1.19" } });
    const { panel, client, reload, store } = setup([info(), running("Installing the new version"), new Error("down"), info({ current: "0.1.20", updateAvailable: false, status: { state: "done", to: "0.1.20" } })]);
    await panel.refresh();
    button("update-run").click();
    button("update-run").click();
    await flush();
    expect(client.startUpdate).toHaveBeenCalledOnce();
    expect(store.data.get("excelPartner.updating")).toBe("0.1.19");
    expect(button("update-check").disabled).toBe(true);
    expect(el("update-run").hidden).toBe(true);

    await vi.advanceTimersByTimeAsync(100);
    expect(el("update-status").textContent).toBe("Installing the new version… Keep Excel open.");
    await vi.advanceTimersByTimeAsync(100);
    expect(el("update-status").textContent).toBe("Restarting the connection… Keep Excel open.");
    await vi.advanceTimersByTimeAsync(100);
    expect(el("update-status").textContent).toBe("Updated to 0.1.20. Reloading…");
    expect(el("update-status").dataset.kind).toBe("ok");
    expect(store.data.get("excelPartner.updating")).toBe("0.1.19"); // kept: the reloaded pane announces the result from it
    expect(reload).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(1300);
    expect(reload).toHaveBeenCalledOnce();
  });

  it("shows the error and frees the buttons when the update fails", async () => {
    const failed = info({ updateAvailable: false, status: { state: "failed", error: "npm exited with code 1" } });
    const { panel, reload, store } = setup([info(), failed]);
    await panel.refresh();
    button("update-run").click();
    button("update-run").click();
    await flush();
    await vi.advanceTimersByTimeAsync(100);
    expect(el("update-status").textContent).toContain("did not finish: npm exited with code 1");
    expect(el("update-status").dataset.kind).toBe("error");
    expect(button("update-check").disabled).toBe(false);
    expect(store.data.has("excelPartner.updating")).toBe(false);
    await vi.advanceTimersByTimeAsync(5000);
    expect(reload).not.toHaveBeenCalled();
  });

  it("stops waiting when the bridge never comes back", async () => {
    const { panel } = setup([info(), new Error("down")]);
    await panel.refresh();
    button("update-run").click();
    button("update-run").click();
    await flush();
    await vi.advanceTimersByTimeAsync(1200);
    expect(el("update-status").textContent).toContain("did not come back");
    expect(button("update-check").disabled).toBe(false);
  });

  it("does not wait forever for an update that never started", async () => {
    const { panel } = setup([info()]);
    await panel.refresh();
    button("update-run").click();
    button("update-run").click();
    await flush();
    await vi.advanceTimersByTimeAsync(31_000);
    expect(el("update-status").textContent).toBe("The update did not start. Try again.");
  });

  it("shows why the bridge refused to start, and does not follow anything", async () => {
    const { panel, client } = setup([info()], { startUpdate: vi.fn(async () => { throw new Error("The assistant is still working."); }) });
    await panel.refresh();
    button("update-run").click();
    button("update-run").click();
    await flush();
    expect(el("update-status").textContent).toBe("Could not start the update: The assistant is still working.");
    expect(el("update-status").dataset.kind).toBe("error");
    await vi.advanceTimersByTimeAsync(500);
    expect(client.updateInfo).toHaveBeenCalledTimes(1);
    expect(button("update-run").disabled).toBe(false);
  });

  it("picks an update that is already running back up when Settings opens", async () => {
    const { panel } = setup([info({ status: { state: "running", message: "Stopping the bridge", from: "0.1.19" } })]);
    await panel.refresh();
    expect(el("update-status").textContent).toBe("Stopping the bridge… Keep Excel open.");
    expect(button("update-check").disabled).toBe(true);
    expect(el("update-run").hidden).toBe(true);
  });
});

describe("after the pane reloads", () => {
  it("says how an update that was started here ended, once", async () => {
    const { panel, announce, store } = setup([info({ current: "0.1.20", updateAvailable: false, status: { state: "done", to: "0.1.20" } })]);
    store.setItem("excelPartner.updating", "0.1.19");
    await panel.afterReload();
    expect(announce).toHaveBeenCalledWith("Excel Partner was updated to version 0.1.20.");
    expect(store.data.has("excelPartner.updating")).toBe(false);
    await panel.afterReload();
    expect(announce).toHaveBeenCalledOnce();
  });

  it("reports a failure too", async () => {
    const { panel, announce, store } = setup([info({ status: { state: "failed", error: "EBUSY" } })]);
    store.setItem("excelPartner.updating", "0.1.19");
    await panel.afterReload();
    expect(announce).toHaveBeenCalledWith(expect.stringContaining("did not finish: EBUSY"));
  });

  it("says nothing when this pane did not start an update", async () => {
    const { panel, announce, client } = setup([info({ status: { state: "done", to: "0.1.20" } })]);
    await panel.afterReload();
    expect(client.updateInfo).not.toHaveBeenCalled();
    expect(announce).not.toHaveBeenCalled();
  });

  it("waits while the update is still running, and keeps the marker if the bridge is not up yet", async () => {
    const running = setup([info({ status: { state: "running" } })]);
    running.store.setItem("excelPartner.updating", "0.1.19");
    await running.panel.afterReload();
    expect(running.announce).not.toHaveBeenCalled();
    expect(running.store.data.has("excelPartner.updating")).toBe(true);
    const down = setup([new Error("down")]);
    down.store.setItem("excelPartner.updating", "0.1.19");
    await down.panel.afterReload();
    expect(down.store.data.has("excelPartner.updating")).toBe(true);
  });
});
