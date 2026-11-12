import type http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { z } from "zod";
import { ProviderIdSchema } from "../../../protocol/src/index.js";
import { providerPackageRoot } from "../../../providers/src/profiled.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { SessionManager } from "../session-manager.js";
import { UpdateError, type UpdateService } from "../updater.js";
import type { ProviderRegistry } from "../../../providers/src/registry.js";

const root = providerPackageRoot();
const store = createRequire(import.meta.url)(path.join(root, "bin", "runtime-store.cjs"));
const action = z.object({ action: z.enum(["install", "remove", "repair", "debug-on", "debug-off"]), provider: ProviderIdSchema.optional() });
export class SystemRoutes {
  constructor(private readonly sessions: SessionManager, private readonly providers: ProviderRegistry, private readonly updater: UpdateService | null,
    private readonly loginBusy = () => false) {}
  async handle(req: http.IncomingMessage, res: http.ServerResponse, url: URL) {
    const session = connectedSession(this.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const dataDir = store.home();
    try {
      const manager = await import(pathToFileURL(path.join(root, "bin", "runtime-manager.mjs")).href);
      if (req.method === "GET") {
        if (url.searchParams.get("diagnostics") === "1") {
          const doctor = await import(pathToFileURL(path.join(root, "bin", "doctor-checks.mjs")).href);
          const provider = ProviderIdSchema.safeParse(url.searchParams.get("provider"));
          return sendJson(res, 200, await doctor.collectDoctor({ root, provider: provider.success ? provider.data : undefined,
            probe: url.searchParams.get("probe") === "1", mcp: url.searchParams.get("mcp") === "1", host: session.host }));
        }
        const deployment = store.currentDeployment(dataDir);
        let debug = false;
        try { debug = JSON.parse(fs.readFileSync(path.join(dataDir, "debug.json"), "utf8")).expires > Date.now(); } catch { /* No debug window is open. */ }
        return sendJson(res, 200, { debug, version: deployment.releases[deployment.active]?.version ?? null,
          providers: manager.runtimeInventory(dataDir) });
      }
      if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
      const input = action.parse(JSON.parse(await readBody(req)));
      if (this.loginBusy() || this.sessions.list().some(item => item.isBusy)) return sendJson(res, 409, { error: "Wait for the current response or sign-in to finish." });
      if (input.action === "repair") {
        if (!this.updater) return sendJson(res, 503, { error: "Recovery controls are unavailable in this installation." });
        return sendJson(res, 202, { status: this.updater.start({ busy: false, action: "repair" }) });
      }
      if (input.action.startsWith("debug-")) {
        store.atomic(path.join(dataDir, "debug.json"), { expires: input.action === "debug-on" ? Date.now() + 15 * 60000 : 0 });
        return sendJson(res, 200, { ok: true });
      }
      if (!input.provider) return sendJson(res, 400, { error: "Choose a provider." });
      if (input.action === "remove") {
        store.changeState(dataDir, (state: any) => { state.selected = state.selected.filter((id: string) => id !== input.provider); delete state.active[input.provider!]; delete state.pending?.[input.provider!]; state.disabled = [...new Set([...(state.disabled ?? []), input.provider])]; });
      } else {
        const deployment = await import(pathToFileURL(path.join(root, "bin", "manage.mjs")).href);
        const baseline = deployment.ensureBaseline(root, dataDir);
        await manager.installRuntime(root, input.provider, { dataDir, adapterRoot: baseline.root });
      }
      this.providers.invalidateModels(input.provider);
      return sendJson(res, 200, { ok: true });
    } catch (error) {
      if (error instanceof UpdateError) return sendJson(res, 409, { error: error.message, code: error.code });
      const message = String(error);
      const code = message.includes("NATIVE_PREFLIGHT_AUTH_REQUIRED") ? "login_required" : message.includes("OPERATION_LOCKED") ? "operation_locked" : "system_action_failed";
      return sendJson(res, 400, { code, error: code === "login_required" ? "Private CLI installed. Sign in using this provider, then retry installation. Your previous CLI is still selected."
        : "The operation did not finish. Your previous configuration is retained; export diagnostics before retrying." });
    }
  }
}
