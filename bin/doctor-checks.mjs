import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createHash, randomUUID, X509Certificate } from "node:crypto";
import { createRequire } from "node:module";
import { runCommand, requestHealth, probeMcp, readJson, readLastStartup } from "./doctor-probes.mjs";

export const PROVIDERS = ["claude", "copilot", "codex", "opencode"];
const version = text => /\b(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)\b/.exec(String(text))?.[1] ?? null;
const older = (a, b) => {
  const left = a.split(/[.-]/).slice(0, 3).map(Number), right = b.split(/[.-]/).slice(0, 3).map(Number);
  for (let i = 0; i < 3; i++) if (left[i] !== right[i]) return left[i] < right[i];
  return false;
};
const hash = text => createHash("sha256").update(text).digest("hex").slice(0, 16);
const file = candidate => { try { return fs.statSync(candidate).isFile(); } catch { return false; } };
function installedPackage(root, name) {
  try {
    let directory = path.dirname(createRequire(path.join(root, "package.json")).resolve(name));
    for (let depth = 0; depth < 8; depth++, directory = path.dirname(directory)) {
      try { const pkg = readJson(path.join(directory, "package.json")); if (pkg.name === name) return pkg; } catch { /* next */ }
    }
  } catch { /* damaged/missing package */ }
  try { return readJson(path.join(root, "node_modules", name, "package.json")); } catch { return {}; }
}

/** Read PATH in the same case-insensitive environment used by Windows. */
export function envValue(env, name, platform = process.platform) {
  const key = Object.keys(env).find(key => platform === "win32" ? key.toUpperCase() === name.toUpperCase() : key === name);
  return key ? env[key] : undefined;
}

export function resolveCommands(command, env, platform = process.platform) {
  const p = platform === "win32" ? path.win32 : path.posix;
  const extensions = platform === "win32" && !p.extname(command)
    ? (envValue(env, "PATHEXT", platform) ?? ".EXE;.COM;.CMD;.BAT").split(";").filter(ext => /^\.(exe|com|cmd|bat)$/i.test(ext)) : [""];
  const paths = p.isAbsolute(command) || /[\\/]/.test(command) ? [command]
    : (envValue(env, "PATH", platform) ?? "").split(platform === "win32" ? ";" : ":")
      .filter(Boolean).flatMap(dir => extensions.map(extension => p.join(dir.replace(/^"|"$/g, ""), command + extension)));
  return [...new Set(paths.filter(candidate => {
    if (!file(candidate)) return false;
    if (platform !== "win32") { try { fs.accessSync(candidate, fs.constants.X_OK); } catch { return false; } }
    return true;
  }).flatMap(candidate => { try { return [fs.realpathSync(candidate)]; } catch { return []; } }))];
}

function adapterMinimum(root, provider) {
  for (const candidate of [path.join(root, "dist", "providers", "src", `${provider}.js`), path.join(root, "providers", "src", `${provider}.ts`)]) {
    try { return /MIN_VERSION\s*=\s*["']([0-9.]+)["']/.exec(fs.readFileSync(candidate, "utf8"))?.[1] ?? null; } catch { /* next */ }
  }
  return null;
}

function expectedProtocol(root) {
  for (const candidate of ["dist/protocol/src/version.js", "protocol/src/version.ts"]) {
    try { const match = /PROTOCOL_VERSION\s*=\s*(\d+)/.exec(fs.readFileSync(path.join(root, candidate), "utf8")); if (match) return Number(match[1]); } catch { /* next */ }
  }
  return null;
}

// All emitted strings are product-owned. Never include raw process/config/error output.
export async function collectDoctor(options) {
  const { root: requestedRoot, target = "all", provider, probe = false, mcp = false, integrity = false, host = "excel",
    env = process.env, platform = process.platform, arch = process.arch, nodeVersion = process.versions.node,
    run = runCommand, health = requestHealth, checkMcp = probeMcp, now = Date.now } = options;
  const checks = [];
  const add = (id, status, code, message, action, details) => checks.push({ id, status, code, message,
    ...(action ? { action } : {}), ...(details ? { details } : {}) });
  let root = requestedRoot, runtimeStore, runtimeState;
  const dataDir = env.OFFICE_PARTNER_DATA_DIR || path.join(env.LOCALAPPDATA || os.homedir(), "Excel Partner");
  if (file(path.join(root, "bin", "runtime-store.cjs"))) {
    try {
      runtimeStore = createRequire(import.meta.url)(path.join(root, "bin", "runtime-store.cjs"));
      const deployment = runtimeStore.currentDeployment(dataDir), active = deployment.releases[deployment.active];
      if (active) root = active.root;
      runtimeState = runtimeStore.readState(dataDir);
      add("deployment.state", "pass", "DEPLOYMENT_STATE_OK", "Deployment and runtime selection metadata parsed.", undefined,
        { active: active?.version ?? null });
    } catch { add("deployment.state", "fail", "DEPLOYMENT_STATE_INVALID", "Saved deployment metadata cannot be read safely.", "Use the recovery launcher to restore a retained release."); }
  }
  if (integrity && runtimeStore && runtimeState) {
    try {
      const { verifyRelease } = await import("./deployment.mjs");
      const { verifyRuntime } = await import("./runtime-manager.mjs");
      const deployment = runtimeStore.currentDeployment(dataDir);
      for (const release of Object.values(deployment.releases)) {
        try { verifyRelease(release, true); add(`integrity.release.${release.id.slice(0, 16)}`, "pass", "RELEASE_INTEGRITY_OK", "Retained release file inventory matches.", undefined, { version: version(release.version) }); }
        catch { add(`integrity.release.${release.id.slice(0, 16)}`, "fail", "RELEASE_INTEGRITY_FAILED", "A retained release has missing or changed files.", "Use recovery repair or restore a verified cached release.", { version: version(release.version) }); }
      }
      for (const profile of Object.values(runtimeState.profiles)) {
        if (profile.mode !== "managed" || provider && provider !== profile.provider) continue;
        try { verifyRuntime(profile, dataDir); add(`integrity.runtime.${profile.id.slice(0, 16)}`, "pass", "RUNTIME_INTEGRITY_OK", "Managed CLI file inventory matches.", undefined, { provider: profile.provider, version: version(profile.version) }); }
        catch { add(`integrity.runtime.${profile.id.slice(0, 16)}`, "fail", "RUNTIME_INTEGRITY_FAILED", "A managed CLI has missing or changed files.", "office-partner repair", { provider: profile.provider, version: version(profile.version) }); }
      }
      add("integrity.scan", "pass", "INTEGRITY_SCAN_COMPLETE", "Retained app and selected provider inventories were checked. No CLI was launched.");
    } catch { add("integrity.scan", "fail", "INTEGRITY_TOOLS_UNAVAILABLE", "Integrity checks could not run.", "Use the retained recovery launcher or reinstall the management tools."); }
  } else if (integrity) add("integrity.scan", runtimeState ? "skip" : "warn", "INTEGRITY_STATE_UNAVAILABLE", "No readable managed inventory is available for integrity checks.");
  let pkg = {};
  try { pkg = readJson(path.join(root, "package.json")); } catch { /* report below */ }
  const command = pkg.name === "excel-partner" ? "excel-partner" : "office-partner";
  const packageName = ["excel-partner", "office-partner"].includes(pkg.name) ? pkg.name : "unknown";
  add("installation.package", packageName !== "unknown" && version(pkg.version) ? "pass" : "fail", packageName !== "unknown" && version(pkg.version) ? "PACKAGE_OK" : "PACKAGE_INVALID",
    packageName !== "unknown" ? "Product package identified." : "Product package metadata is missing or invalid.", `Reinstall the exact product release if package metadata is damaged.`);
  const restricted = typeof pkg.engines?.node === "string" && pkg.engines.node.includes("<26.0.0");
  const nodeSupported = restricted ? !older(nodeVersion, "20.19.0") && (Number(nodeVersion.split(".")[0]) === 20 || !older(nodeVersion, "22.12.0")) && older(nodeVersion, "26.0.0") : !older(nodeVersion, "20.0.0");
  add("installation.node", nodeSupported ? "pass" : "fail", nodeSupported ? "NODE_OK" : "NODE_UNSUPPORTED",
    "Check Node.js against the product and packaged SQLite driver requirements.", "Use a Node.js version supported by this release.", { version: version(nodeVersion), minimum: restricted ? "20.19.0" : "20.0.0", ...(restricted ? { supported: "20.19+ within major 20, or 22.12+ below 26" } : {}) });

  let sdk = {};
  sdk = installedPackage(root, "@github/copilot-sdk");
  const sdkVersion = version(sdk.version), sdkRuntime = version(sdk.copilotCliVersion);
  if (!provider || provider === "copilot") {
    add("installation.copilot_sdk", sdkVersion ? "pass" : provider === "copilot" ? "fail" : "warn", sdkVersion ? "SDK_FOUND" : "SDK_MISSING",
      sdkVersion ? "Copilot SDK metadata found; its bundled runtime is not necessarily the CLI selected by this release." : "Copilot SDK dependency is missing or unreadable.",
      sdkVersion ? undefined : "Repair dependencies for the installed product version.", { version: sdkVersion, bundledRuntime: sdkRuntime });
    if (sdk.engines?.node === "^20.19.0 || >=22.12.0") {
      const major = Number(nodeVersion.split(".")[0]);
      const supported = major === 20 && !older(nodeVersion, "20.19.0") || !older(nodeVersion, "22.12.0");
      add("installation.sdk_node", supported ? "pass" : "fail", supported ? "SDK_NODE_OK" : "SDK_NODE_UNSUPPORTED",
        "Check Node.js against the installed Copilot SDK requirement.", "Use Node 20.19+ within major 20, or Node 22.12+.");
    }
  }

  const addinData = dataDir;
  const gatewayData = env.AGENT_GATEWAY_HOME ?? path.join(env.LOCALAPPDATA || path.join(os.homedir(), ".config"), "Agent Gateway");
  const gatewayConfigFile = path.join(gatewayData, "config.json");
  let gatewayConfig = {}, gatewayConfigValid = true;
  if (file(gatewayConfigFile)) {
    try {
      gatewayConfig = readJson(gatewayConfigFile);
      gatewayConfigValid = gatewayConfig && typeof gatewayConfig === "object" && !Array.isArray(gatewayConfig)
        && (gatewayConfig.port === undefined || Number.isInteger(gatewayConfig.port) && gatewayConfig.port > 0 && gatewayConfig.port <= 65535)
        && (gatewayConfig.enabledProviders === undefined || Array.isArray(gatewayConfig.enabledProviders) && gatewayConfig.enabledProviders.length > 0
          && gatewayConfig.enabledProviders.every(id => PROVIDERS.includes(id)));
    } catch { gatewayConfigValid = false; }
  }
  if (!gatewayConfigValid) gatewayConfig = {};
  const services = [];
  if (target !== "gateway") services.push({ id: "addin", dir: addinData, cert: env.BRIDGE_CERT ?? path.join(addinData, "localhost.crt"),
    key: env.BRIDGE_KEY ?? path.join(addinData, "localhost.key"), port: Number(env.BRIDGE_PORT ?? 8765), entry: "dist/bridge/src/index.js", cli: `${command} setup` });
  if (target !== "addin") services.push({ id: "gateway", dir: gatewayData, cert: path.join(gatewayData, "localhost.crt"),
    key: path.join(gatewayData, "localhost.key"), port: gatewayConfig.port ?? 8766, entry: "gateway/dist/server.mjs", cli: `${command} setup gateway` });
  if (target !== "addin") add("gateway.config", gatewayConfigValid ? "pass" : "fail", gatewayConfigValid ? "CONFIG_OK" : "CONFIG_INVALID",
    gatewayConfigValid ? "Gateway configuration syntax and provider/port settings checked." : "Gateway configuration is invalid.", "Restore a valid gateway configuration; credentials were not read or exported.");
  if (target !== "gateway") {
    let valid = true;
    const config = env.BRIDGE_CONFIG ?? path.join(addinData, "config.json");
    if (file(config)) { try { const value = readJson(config); valid = Boolean(value && typeof value === "object" && !Array.isArray(value)); } catch { valid = false; } }
    add("addin.config", valid ? "pass" : "fail", valid ? "CONFIG_OK" : "CONFIG_INVALID",
      valid ? "Local add-in configuration syntax checked (defaults if absent)." : "Local add-in configuration is invalid.", "Restore valid local configuration; its contents are not included in this report.");
  }
  for (const service of services) {
    const present = file(path.join(root, service.entry));
    add(`${service.id}.build`, present ? "pass" : "fail", present ? "BUILD_OK" : "BUILD_MISSING", present ? "Compiled service entrypoint is present." : "Compiled service entrypoint is missing.", "Reinstall this product version, or build the source checkout.");
    if (platform === "win32") {
      let startup;
      try { startup = await run("reg.exe", ["query", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "/v", service.id === "addin" ? "Excel Partner Bridge" : "Agent Gateway"], { env, timeoutMs: 3000 }); }
      catch { startup = { code: null }; }
      add(`${service.id}.startup`, startup.code === 0 ? "pass" : "warn", startup.code === 0 ? "STARTUP_OK" : "STARTUP_UNVERIFIED",
        startup.code === 0 ? "Windows sign-in startup entry is present." : "Windows sign-in startup entry was not verified.", service.cli);
      // The entry stores the full path of Node and of the launcher. If either was moved or removed, nothing starts at sign-in.
      const programs = startup.code === 0 && typeof startup.stdout === "string" ? [...startup.stdout.matchAll(/"([^"]+)"/g)].map(match => match[1]).slice(0, 2) : [];
      if (programs.length === 2) {
        const present = programs.every(program => file(program));
        add(`${service.id}.startup_target`, present ? "pass" : "fail", present ? "STARTUP_TARGET_OK" : "STARTUP_TARGET_MISSING",
          present ? "The sign-in startup entry points to files that exist." : "The sign-in startup entry points to Node.js or a launcher file that no longer exists.", present ? undefined : service.cli);
      }
      const attempt = readLastStartup(dataDir, service.id);
      if (attempt) add(`${service.id}.startup_last`, attempt.ok ? "pass" : "warn", attempt.ok ? "STARTUP_LAST_OK" : "STARTUP_LAST_FAILED",
        attempt.ok ? "The most recent start at sign-in or from Fix connection succeeded." : "The most recent start at sign-in or from Fix connection failed.",
        attempt.ok ? undefined : 'Open Start > "Office Partner - Fix connection".', { at: String(attempt.at ?? ""), attempt: Number(attempt.attempt) || 1, ...(attempt.ok ? {} : { reason: String(attempt.code ?? "START_FAILED") }) });
    }
    let ca;
    if (!file(service.cert) || !file(service.key)) {
      add(`${service.id}.tls`, "fail", "TLS_MISSING", `Not set up: certificate or key missing. Run: ${service.cli}`, service.cli);
    } else {
      try {
        ca = fs.readFileSync(service.cert);
        const cert = new X509Certificate(ca), expires = Date.parse(cert.validTo), starts = Date.parse(cert.validFrom);
        const valid = expires > now() && starts <= now();
        add(`${service.id}.tls`, !valid ? "fail" : expires - now() < 30 * 86400000 ? "warn" : "pass",
          !valid ? "TLS_INVALID_DATE" : expires - now() < 30 * 86400000 ? "TLS_EXPIRING" : "TLS_OK",
          valid ? "Local certificate parsed and validity checked." : "Local certificate is expired or not yet valid.",
          "Check the system clock and local certificate registration.", { daysRemaining: Math.floor((expires - now()) / 86400000) });
      } catch { add(`${service.id}.tls`, "fail", "TLS_INVALID", "Local certificate cannot be parsed.", "Restore the local certificate before restarting."); }
    }
    if (!Number.isInteger(service.port) || service.port < 1 || service.port > 65535) {
      add(`${service.id}.health`, "fail", "PORT_INVALID", "Configured local service port is invalid.", "Correct the local port configuration."); continue;
    }
    if (!ca) { add(`${service.id}.health`, "skip", "HEALTH_NOT_CHECKED", "Trusted HTTPS health check requires a readable local certificate."); continue; }
    let result;
    try { result = await health({ port: service.port, ca }); } catch { result = { errorCode: "REQUEST_FAILED" }; }
    const error = result.errorCode;
    const code = error ? error === "ECONNREFUSED" ? "SERVICE_STOPPED" : error === "ETIMEDOUT" ? "SERVICE_TIMEOUT"
      : /CERT|TLS|SSL|VERIFY/.test(error) ? "SERVICE_TLS_ERROR" : "SERVICE_UNREACHABLE"
      : !result[service.id === "addin" ? "bridge" : "gateway"] ? "SERVICE_IDENTITY_MISMATCH"
        : service.id === "gateway" && !result.ready ? "SERVICE_NOT_READY" : "SERVICE_OK";
    const label = `${service.id === "addin" ? "Office Partner bridge" : "Gateway"} on port ${service.port}`;
    const healthMessage = code === "SERVICE_OK" ? `${label} answered its health check and is ready.`
      : code === "SERVICE_NOT_READY" ? `${label} is responding but still initializing its model catalog or warmup. Run the check again after startup finishes.`
      : code === "SERVICE_STOPPED" ? `${label} is not running or is not listening on this port.`
      : code === "SERVICE_IDENTITY_MISMATCH" ? `${label}: a different service answered on this port.`
      : code === "SERVICE_TIMEOUT" ? `${label} did not answer within the health-check timeout.`
      : code === "SERVICE_TLS_ERROR" ? `${label} could not be verified using its local HTTPS certificate.`
      : `${label} could not be reached or returned an invalid health response.`;
    add(`${service.id}.health`, code === "SERVICE_OK" ? "pass" : "fail", code, healthMessage,
      code === "SERVICE_NOT_READY" ? "Wait for gateway initialization to finish, then run the check again."
        : code === "SERVICE_IDENTITY_MISMATCH" ? "Check the port owner; do not terminate an unrelated process." : `${command} start ${service.id === "addin" ? "addin" : "gateway"}`,
      { port: service.port, ...(service.id === "addin" && !error ? { protocolVersion: result.protocolVersion } : {}) });
    if (service.id === "addin" && code === "SERVICE_OK") {
      const expected = expectedProtocol(root), actual = result.protocolVersion ?? null;
      add("addin.protocol", expected === null || actual === null ? "warn" : expected === actual ? "pass" : "fail",
        expected === null || actual === null ? "PROTOCOL_UNVERIFIED" : expected === actual ? "PROTOCOL_OK" : "PROTOCOL_MISMATCH",
        "Compare the running bridge protocol to this product's protocol.", "Restart the bridge from the intended product installation and reload the pane.", { expected, actual });
    }
  }

  for (const id of provider ? [provider] : PROVIDERS) {
    const failed = runtimeState?.failures?.[id];
    if (failed && /^(NATIVE_PREFLIGHT_[A-Z_]+|INSTALL_FAILED|INSTALL_TIMEOUT|CLI_PREFLIGHT_FAILED|RUNTIME_INSTALL_FAILED|RUNTIME_INTEGRITY_FAILED|ACTIVATION_CANCELLED)$/.test(failed.code)) {
      add(`provider.${id}.activation`, "warn", failed.code, `${id}: the private runtime has not been activated; its previous selection was retained.`,
        failed.code === "NATIVE_PREFLIGHT_AUTH_REQUIRED" ? `office-partner providers login ${id}` : `office-partner providers install ${id}`, { version: version(failed.version) });
    }
    let profile;
    try { profile = runtimeStore?.selectProfile(dataDir, id); }
    catch (error) {
      const disabled = error.message === "PROVIDER_DISABLED";
      add(`provider.${id}.profile`, disabled ? "skip" : "fail", disabled ? "PROVIDER_DISABLED" : "RUNTIME_PROFILE_UNAVAILABLE",
        disabled ? `${id}: provider is disabled.` : `${id}: recorded runtime is missing, changed, or invalid.`, `office-partner providers install ${id}`); continue;
    }
    if (profile?.fallbackFrom) add(`provider.${id}.fallback`, "warn", "RUNTIME_FALLBACK_ACTIVE", `${id}: a retained CLI is being used because the selected entrypoint is missing or changed.`, "office-partner repair");
    const override = envValue(env, `${id.toUpperCase()}_BIN`, platform);
    const matches = resolveCommands(profile?.command ?? override ?? id, env, platform), minimum = adapterMinimum(profile?.adapterRoot ?? root, id);
    const required = Boolean(provider) || target !== "addin" && (gatewayConfig.enabledProviders ?? ["copilot"]).includes(id);
    if (!matches.length) {
      add(`provider.${id}.executable`, required ? "fail" : "warn", "CLI_NOT_FOUND", `${id}: CLI executable not found.`, "Check the CLI installation and executable override in the service environment.", { source: override ? "override" : "PATH" });
      continue;
    }
    add(`provider.${id}.executable`, matches.length > 1 ? "warn" : "pass", matches.length > 1 ? "CLI_MULTIPLE" : "CLI_FOUND",
      `${id}: ${matches.length > 1 ? "multiple executable candidates found" : "executable found"}.`, matches.length > 1 ? "Check which installation appears first on PATH." : undefined,
      { source: profile ? profile.fallbackFrom ? "fallback" : profile.mode : override ? "override" : "PATH", candidates: matches.length, executableHash: hash(matches[0]), extension: [".exe", ".cmd", ".bat", ".com", ".js", ".mjs"].includes(path.extname(matches[0]).toLowerCase()) ? path.extname(matches[0]).toLowerCase() : "other", minimum });
    if (!probe) {
      add(`provider.${id}.probe`, "skip", "CLI_NOT_PROBED", `${id}: version and authentication have not been verified.`, `${command} doctor --provider ${id} --probe`); continue;
    }
    const childEnv = runtimeStore ? { ...runtimeStore.childEnvironment(env, profile), NO_COLOR: "1" }
      : { ...env, COPILOT_AUTO_UPDATE: "false", DISABLE_AUTOUPDATER: "1", DISABLE_UPDATES: "1", OPENCODE_DISABLE_AUTOUPDATE: "true", NO_COLOR: "1" };
    let result;
    try { result = await run(matches[0], ["--version"], { env: childEnv }); } catch { result = { code: null, errorCode: "SPAWN_FAILED" }; }
    const foundVersion = result.code === 0 ? version(result.stdout) ?? version(result.stderr) : null;
    if (!foundVersion) {
      const code = result.errorCode === "ETIMEDOUT" ? "CLI_TIMEOUT" : result.errorCode === "RUNNER_MISSING" ? "CLI_RUNNER_MISSING" : result.code === 0 ? "CLI_VERSION_UNKNOWN" : "CLI_SPAWN_FAILED";
      add(`provider.${id}.version`, "fail", code, `${id}: could not verify the CLI version.`, "Check the CLI installation and command shim; raw CLI output is intentionally omitted.", { exitCode: Number.isInteger(result.code) ? result.code : null }); continue;
    }
    add(`provider.${id}.version`, minimum && older(foundVersion, minimum) ? "fail" : "pass", minimum && older(foundVersion, minimum) ? "CLI_TOO_OLD" : "CLI_VERSION_OK",
      `${id}: CLI version parsed; newer versions still need compatibility validation.`, minimum && older(foundVersion, minimum) ? "Use a CLI version supported by this product release." : undefined, { version: foundVersion, minimum });
    if (id === "copilot") {
      if (sdkRuntime && sdkRuntime !== foundVersion) add(`provider.${id}.compatibility`, "warn", "SDK_CLI_PAIR_UNVERIFIED",
        "External Copilot CLI differs from the SDK's bundled runtime. Compatibility is unverified.",
        "Include both version numbers when reporting a Copilot tool failure.", { cliVersion: foundVersion, sdkVersion, bundledRuntime: sdkRuntime });
      add(`provider.${id}.auth`, "warn", "AUTH_UNVERIFIED", "Copilot authentication cannot be proved by CLI presence or a stored config file.", "Use the product sign-in flow. A successful tool turn is a separate live check."); continue;
    }
    const args = id === "claude" ? ["auth", "status"] : id === "codex" ? ["login", "status"] : ["auth", "list"];
    try { result = await run(matches[0], args, { env: childEnv }); } catch { result = { code: null }; }
    let loggedIn = null;
    const authText = `${result.stdout ?? ""}\n${result.stderr ?? ""}`.replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, "");
    if (result.code === 0) {
      if (id === "claude") { try { const value = JSON.parse(result.stdout)?.loggedIn; if (typeof value === "boolean") loggedIn = value; } catch { /* unknown */ } }
      else if (id === "codex") loggedIn = /not logged in|logged out/i.test(authText) ? false : /(?:^|\n)\s*logged in\b/i.test(authText) ? true : null;
      else { const count = /\b(\d+)\s+credentials?\b/i.exec(authText); if (count) loggedIn = Number(count[1]) > 0; }
    }
    if (id === "codex" && /not logged in|logged out/i.test(authText)) loggedIn = false;
    add(`provider.${id}.auth`, loggedIn ? "pass" : loggedIn === false && required ? "fail" : "warn",
      loggedIn ? "AUTH_PRESENT" : loggedIn === false ? "AUTH_REQUIRED" : result.errorCode === "ETIMEDOUT" ? "AUTH_TIMEOUT" : "AUTH_UNVERIFIED",
      loggedIn ? `${id}: CLI reports an existing sign-in; model access is not verified.` : `${id}: sign-in was not verified.`,
      loggedIn ? undefined : "Check sign-in using this provider's own CLI. No login was started by doctor.");
  }
  if (mcp) {
    try {
      const result = await checkMcp({ root, host });
      add("mcp.catalog", "pass", "MCP_CATALOG_OK", "Local MCP initialized and listed tool schemas. No tool was executed.", undefined,
        { host, toolCount: result.count, catalogHash: result.catalogHash });
    } catch (error) {
      const code = error?.message === "ETIMEDOUT" ? "MCP_TIMEOUT" : error?.message === "MCP_ENTRYPOINT_MISSING" ? "MCP_ENTRYPOINT_MISSING"
        : error?.code === "ERR_MODULE_NOT_FOUND" ? "MCP_DEPENDENCY_MISSING"
        : ["CATALOG_INVALID", "CATALOG_EMPTY", "CATALOG_LIMIT", "CATALOG_PAGINATION", "CATALOG_HOST_MISMATCH"].includes(error?.message) ? `MCP_${error.message}` : "MCP_CATALOG_FAILED";
      add("mcp.catalog", "fail", code, "Could not initialize or validate the local MCP catalog.", "Check the compiled MCP entrypoint and its dependencies. This check does not test Office execution.", { host });
    }
  } else add("mcp.catalog", "skip", "MCP_NOT_PROBED", "MCP transport and schemas have not been checked.", `${command} doctor --mcp --host ${host}`);
  const status = checks.some(check => check.status === "fail") ? "fail" : checks.some(check => check.status === "warn") ? "warn" : "pass";
  return { schemaVersion: 1, diagnosticId: randomUUID(), generatedAt: new Date(now()).toISOString(),
    product: { name: packageName, version: version(pkg.version) }, environment: { platform, arch, node: version(nodeVersion) },
    scope: { target, provider: provider ?? null, probe, mcp, integrity, host }, status, checks };
}
