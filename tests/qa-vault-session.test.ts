import { afterAll, beforeAll, describe, expect, it } from "vitest";
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { clearCredentials, dpapiAvailable, envNames, hasCredentials, loadCredentials, storeCredentials, vaultFile } from "../qa/src/vault.js";
import { authFile, checkSession } from "../qa/src/session.js";
import { saveConfig, parseConfig } from "../qa/src/config.js";
import { makeProject } from "./qa-helpers.js";

describe("QA vault", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "qa-vault-"));
  const saved = process.env.QA_VAULT_DIR;
  beforeAll(() => { process.env.QA_VAULT_DIR = dir; });
  afterAll(() => { if (saved === undefined) delete process.env.QA_VAULT_DIR; else process.env.QA_VAULT_DIR = saved; });

  it("reads credentials from environment variables first", async () => {
    const names = envNames("qa-admin");
    expect(names).toEqual({ user: "QA_QA_ADMIN_USER", password: "QA_QA_ADMIN_PASSWORD" });
    expect(await loadCredentials("https://a.test", "qa-admin", { [names.user]: "u", [names.password]: "p" })).toEqual({ user: "u", password: "p" });
    expect(await loadCredentials("https://a.test", "nobody", {})).toBeNull();
  });

  it.runIf(dpapiAvailable())("stores a password encrypted for this Windows user and never writes it in clear", async () => {
    const file = await storeCredentials("https://a.test", "admin", { user: "boss@example.test", password: "S3cret!-pässwörd" });
    expect(file).toBe(vaultFile("https://a.test", "admin"));
    const onDisk = fs.readFileSync(file, "utf8");
    expect(onDisk).not.toContain("S3cret");
    expect(onDisk).not.toContain("boss@example.test");
    expect(hasCredentials("https://a.test", "admin")).toBe(true);
    expect(await loadCredentials("https://a.test", "admin", {})).toEqual({ user: "boss@example.test", password: "S3cret!-pässwörd" });
    // A different site cannot read it.
    expect(await loadCredentials("https://other.test", "admin", {})).toBeNull();
    expect(clearCredentials("https://a.test", "admin")).toBe(true);
    expect(hasCredentials("https://a.test", "admin")).toBe(false);
  });

  it.runIf(dpapiAvailable())("reports a damaged secret file instead of returning garbage", async () => {
    const file = vaultFile("https://a.test", "broken");
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, "not-a-dpapi-blob");
    await expect(loadCredentials("https://a.test", "broken", {})).rejects.toThrow(/decrypt/);
  });

  it("requires both a user name and a password", async () => {
    await expect(storeCredentials("https://a.test", "x", { user: "", password: "p" })).rejects.toThrow();
  });
});

describe("QA session check", () => {
  let server: http.Server;
  let origin = "";
  const seen: Array<{ url: string; cookie?: string }> = [];
  let uid: number | false = 7;
  let redirect = false;

  beforeAll(async () => {
    server = http.createServer((req, res) => {
      seen.push({ url: req.url ?? "", cookie: req.headers.cookie });
      if (redirect) { res.writeHead(302, { Location: "/login" }).end(); return; }
      res.setHeader("Content-Type", "application/json");
      res.end(JSON.stringify({ jsonrpc: "2.0", result: { uid } }));
    });
    await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
    origin = `http://127.0.0.1:${(server.address() as { port: number }).port}`;
  });
  afterAll(() => new Promise<void>(resolve => server.close(() => resolve())));

  const setup = (type: "odoo-session" | "http-ok" = "odoo-session") => {
    const { dir } = makeProject("odoo", origin);
    const config = parseConfig({ defaultEnvironment: "staging", environments: { staging: { baseUrl: origin } }, profiles: { admin: { check: { type, path: type === "odoo-session" ? "/web/session/get_session_info" : "/" } } } });
    saveConfig(dir, config);
    const writeState = (cookies: unknown[]) => {
      fs.mkdirSync(path.dirname(authFile(dir, "staging", "admin")), { recursive: true });
      fs.writeFileSync(authFile(dir, "staging", "admin"), JSON.stringify({ cookies, origins: [] }));
    };
    return { dir, config, writeState };
  };
  const cookie = (extra: Record<string, unknown> = {}) => ({ name: "session_id", value: "abc", domain: "127.0.0.1", path: "/", expires: -1, ...extra });

  it("reports a missing session without calling the server", async () => {
    const { dir, config } = setup();
    seen.length = 0;
    expect(await checkSession(dir, config, "staging", "admin")).toMatchObject({ valid: false, reason: "no saved session" });
    expect(seen).toHaveLength(0);
  });

  it("accepts a session the Odoo server recognises and sends the cookie only to the configured host", async () => {
    const { dir, config, writeState } = setup();
    writeState([cookie(), cookie({ name: "other", value: "no", domain: "evil.example.org" })]);
    seen.length = 0; uid = 7;
    expect(await checkSession(dir, config, "staging", "admin")).toMatchObject({ valid: true });
    expect(seen[0]).toMatchObject({ url: "/web/session/get_session_info", cookie: "session_id=abc" });
  });

  it("rejects a session the server no longer knows (uid false)", async () => {
    const { dir, config, writeState } = setup();
    writeState([cookie()]);
    uid = false;
    const state = await checkSession(dir, config, "staging", "admin");
    expect(state.valid).toBe(false);
    expect(state.reason).toContain("no longer recognises");
    uid = 7;
  });

  it("skips expired cookies without asking the server", async () => {
    const { dir, config, writeState } = setup();
    writeState([cookie({ expires: 1_000_000 })]);
    seen.length = 0;
    expect((await checkSession(dir, config, "staging", "admin")).reason).toContain("no live cookie");
    expect(seen).toHaveLength(0);
  });

  it("treats a redirect to the sign-in page as expired for generic sites", async () => {
    const { dir, config, writeState } = setup("http-ok");
    writeState([cookie()]);
    redirect = true;
    expect((await checkSession(dir, config, "staging", "admin")).reason).toContain("sign-in");
    redirect = false;
  });

  it("flags an unreachable server so the run does not try to log in blindly", async () => {
    const { dir, writeState } = setup();
    const config = parseConfig({ defaultEnvironment: "staging", environments: { staging: { baseUrl: "http://127.0.0.1:1" } }, profiles: { admin: {} } });
    writeState([cookie()]);
    const state = await checkSession(dir, config, "staging", "admin", 1500);
    expect(state).toMatchObject({ valid: false, unreachable: true });
  });

  it("reports a damaged session file clearly", async () => {
    const { dir, config } = setup();
    fs.mkdirSync(path.dirname(authFile(dir, "staging", "admin")), { recursive: true });
    fs.writeFileSync(authFile(dir, "staging", "admin"), "{not json");
    await expect(checkSession(dir, config, "staging", "admin")).rejects.toThrow(/not a readable saved session/);
  });
});
