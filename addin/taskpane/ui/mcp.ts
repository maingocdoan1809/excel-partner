import type { BridgeClient } from "../bridge-client.js";
interface Server { id: string; name: string; transport: "stdio" | "http"; target: string; args: string[] }
interface Catalog { servers: Server[]; enabled: string[] }

export function setupMcp(getClient: () => BridgeClient) {
  const section = document.createElement("section");
  section.className = "settings-section mcp-settings";
  section.innerHTML = `<h3>MCP servers</h3><p class="settings-description">Connect tools once and use them with any AI provider. Enable a server in Tools for this file.</p><div class="mcp-server-list"></div><details><summary>Add MCP server</summary><form class="mcp-form"><label>Name<input name="name" required maxlength="100" placeholder="My tools"></label><label>Connection<select name="transport"><option value="stdio">Local command (stdio)</option><option value="http">Remote URL (HTTP)</option></select></label><label>Command or URL<input name="target" required placeholder="npx or https://example.com/mcp"></label><label>Command arguments (JSON array)<textarea name="args" rows="2" placeholder='["--yes", "my-mcp-server@1.0.0"]'></textarea></label><p class="settings-description">Local servers run with your user permissions. Test connection launches the command or contacts the URL. Only add servers you trust. Authentication is not supported in this first version.</p><button type="submit">Save server</button></form></details><p class="mcp-status" role="status" aria-live="polite"></p>`;
  document.getElementById("settings-tools-section")!.parentElement!.append(section);
  const picker = document.createElement("div"); picker.className = "mcp-picker";
  document.getElementById("tools-menu")!.append(picker);
  const selectionKey = () => `officePartner.mcpSelection.${document.body.dataset.host ?? "excel"}`;
  const signature = (server: Server) => JSON.stringify([server.name, server.transport, server.target, server.args]);
  const savedSelection = (): Record<string, string> => {
    try { const value = JSON.parse(sessionStorage.getItem(selectionKey()) ?? "{}"); return value && typeof value === "object" && !Array.isArray(value) ? value : {}; }
    catch { return {}; }
  };
  const storeSelection = (selection: Record<string, string>) => { try { sessionStorage.setItem(selectionKey(), JSON.stringify(selection)); } catch { /* Still usable for this connection. */ } };
  let pending: Promise<unknown> = Promise.resolve();
  const status = section.querySelector<HTMLElement>(".mcp-status")!;
  let editingId: string | undefined;
  const list = section.querySelector<HTMLElement>(".mcp-server-list")!;
  const action = async (run: () => Promise<unknown>) => {
    status.textContent = "Working…";
    try { await run(); await refresh(); }
    catch (error) { status.textContent = error instanceof Error ? error.message : String(error); }
  };
  const button = (text: string, run: () => Promise<unknown>) => {
    const node = document.createElement("button"); node.type = "button"; node.textContent = text; node.className = "secondary";
    node.addEventListener("click", () => { node.disabled = true; void action(run).finally(() => { node.disabled = false; }); }); return node;
  };
  const refresh = async () => {
    const data = await getClient().mcpRequest<Catalog>();
    list.replaceChildren(); picker.replaceChildren();
    const heading = document.createElement("h3"); heading.textContent = "MCP servers"; picker.append(heading);
    if (!data.servers.length) { picker.append("Add a server in Settings to connect more tools."); list.textContent = "No MCP servers connected yet."; }
    for (const server of data.servers) {
      const row = document.createElement("div"); row.className = "mcp-server-row";
      const title = document.createElement("strong"); title.textContent = server.name;
      const detail = document.createElement("small"); detail.textContent = `${server.transport} · ${server.target}`;
      row.append(title, detail,
        button("Edit", async () => {
          editingId = server.id;
          const form = section.querySelector<HTMLFormElement>("form")!;
          for (const name of ["name", "transport", "target", "args"] as const) {
            (form.elements.namedItem(name) as HTMLInputElement).value = name === "args" ? JSON.stringify(server.args) : server[name];
          }
          section.querySelector("details")!.open = true;
          (form.elements.namedItem("name") as HTMLInputElement).focus();
          status.textContent = "Editing server. Save to apply changes; enable it again in Tools afterwards.";
        }),
        button("Test connection", async () => {
          const result = await getClient().mcpRequest<{ tools: { name: string }[] }>("POST", { action: "test", id: server.id });
          status.textContent = `${server.name}: ${result.tools.length} tools available. ${result.tools.map(tool => tool.name).slice(0, 10).join(", ")}`;
        }),
        button("Remove", () => getClient().mcpRequest("DELETE", { id: server.id })));
      list.append(row);
      const label = document.createElement("label"); label.className = "settings-toggle";
      const name = document.createElement("span"); name.textContent = server.name;
      const check = document.createElement("input"); check.type = "checkbox"; check.checked = data.enabled.includes(server.id);
      check.addEventListener("change", async () => {
        check.disabled = true;
        try {
          const enabled = check.checked;
          pending = pending.catch(() => undefined).then(() => getClient().mcpRequest("POST", { action: "select", id: server.id, enabled }));
          await pending;
          const saved = savedSelection();
          if (enabled) saved[server.id] = signature(server); else delete saved[server.id];
          storeSelection(saved);
        }
        catch (error) { check.checked = !check.checked; pending = Promise.resolve(); status.textContent = String(error); }
        finally { check.disabled = false; }
      });
      label.append(name, check); picker.append(label);
    }
  };
  section.querySelector("form")!.addEventListener("submit", event => {
    event.preventDefault(); const form = event.currentTarget as HTMLFormElement; const data = new FormData(form);
    void action(async () => {
      const name = String(data.get("name"));
      const id = editingId ?? `mcp-${crypto.randomUUID().slice(0, 8)}`;
      const args = JSON.parse(String(data.get("args") || "[]"));
      await getClient().mcpRequest("POST", { action: "save", config: { id, name, transport: data.get("transport"), target: data.get("target"), args } });
      form.reset(); editingId = undefined; status.textContent = "Server saved. Test the connection, then enable it in Tools.";
    });
  });
  return {
    refresh: () => refresh().catch(error => { status.textContent = String(error); }),
    ready: () => pending,
    restore: () => {
      pending = (async () => {
        const data = await getClient().mcpRequest<Catalog>();
        const saved = savedSelection(); const retained: Record<string, string> = {};
        for (const server of data.servers) {
          if (saved[server.id] !== signature(server)) continue;
          if (!data.enabled.includes(server.id)) await getClient().mcpRequest("POST", { action: "select", id: server.id, enabled: true });
          retained[server.id] = signature(server);
        }
        storeSelection(retained);
      })().catch(error => { status.textContent = String(error); });
      return pending;
    },
  };
}
