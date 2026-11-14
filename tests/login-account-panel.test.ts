// @vitest-environment jsdom
import fs from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { LoginAccount, ProviderLoginStatus } from "../protocol/src/provider-login.js";
import { setupProviderLogin } from "../frontend/taskpane/ui/provider-login.js";

const idle: ProviderLoginStatus = { state: "idle", provider: "opencode", userCode: null, verificationUri: null, expiresAt: null, message: "Choose the account you want to use." };
const accounts: LoginAccount[] = [
  { id: "zeta", name: "Zeta Cloud", connected: false, methods: [{ index: 0, type: "api", label: "API key", prompts: [] }] },
  { id: "github-copilot", name: "GitHub Copilot", connected: true, methods: [{
    index: 0, type: "oauth", label: "Login with GitHub Copilot", prompts: [
      { key: "deploymentType", type: "select", message: "Select GitHub deployment type", options: [{ label: "GitHub.com", value: "github.com" }, { label: "GitHub Enterprise", value: "enterprise" }] },
      { key: "enterpriseUrl", type: "text", message: "Enter your GitHub Enterprise URL or domain", when: { key: "deploymentType", op: "eq", value: "enterprise" } },
    ],
  }] },
  { id: "openai", name: "OpenAI", connected: false, methods: [
    { index: 0, type: "oauth", label: "ChatGPT Pro/Plus (browser)", prompts: [] },
    { index: 1, type: "api", label: "Manually enter API Key", prompts: [] },
  ] },
];

const el = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

beforeEach(() => {
  document.body.innerHTML = fs.readFileSync("frontend/taskpane/index.html", "utf8").split("<body>")[1]!.split("</body>")[0]!;
});
afterEach(() => { document.body.replaceChildren(); vi.restoreAllMocks(); });

function open() {
  const client = {
    providerLogin: vi.fn(async (method: string): Promise<ProviderLoginStatus> => method === "POST" ? { ...idle, state: "success", message: "Signed in." } : idle),
    loginAccounts: vi.fn(async () => accounts),
    providerLogout: vi.fn(async () => {}),
    submitProviderLoginCode: vi.fn(),
  };
  const panel = setupProviderLogin({ client: () => client as never, provider: () => "opencode", loginKind: () => "account", canLogout: () => true, onSuccess: vi.fn(), isBusy: () => false });
  panel.open();
  return { client, panel };
}

const change = (element: HTMLElement, value?: string) => {
  if (value !== undefined) (element as HTMLInputElement).value = value;
  element.dispatchEvent(new Event(element.tagName === "SELECT" ? "change" : "input", { bubbles: true }));
};

describe("account chooser in the pane", () => {
  it("lists the accounts in the order the provider gives them, with no terminal command or Sign in button", async () => {
    open();
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    expect([...el<HTMLSelectElement>("login-account-provider").options].map(option => option.textContent)).toEqual(["Zeta Cloud", "GitHub Copilot ✓", "OpenAI"]);
    expect(el("login-accounts").hidden).toBe(false);
    expect(el("login-start").hidden).toBe(true);
    expect(el("login-manual").hidden).toBe(true);
  });

  it("filters by name", async () => {
    open();
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    change(el("login-account-search"), "zeta");
    expect([...el<HTMLSelectElement>("login-account-provider").options].map(option => option.value)).toEqual(["zeta"]);
  });

  it("offers the account's methods and asks for a key only for a key method", async () => {
    open();
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    change(el("login-account-provider"), "openai");
    expect(el("login-account-method").hidden).toBe(false);
    expect(el("login-account-key-row").hidden).toBe(true);
    expect(el("login-account-go").textContent).toBe("Continue →");
    change(el("login-account-method"), "1");
    expect(el("login-account-key-row").hidden).toBe(false);
    expect(el("login-account-go").textContent).toBe("Save key →");
  });

  it("asks a follow-up question only while its condition holds, and sends the answers", async () => {
    const { client } = open();
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    change(el("login-account-provider"), "github-copilot");
    const rows = [...el("login-account-prompts").children] as HTMLElement[];
    expect(rows).toHaveLength(2);
    expect(rows[1]!.hidden).toBe(true);
    change(rows[0]!.querySelector("select")!, "enterprise");
    expect(rows[1]!.hidden).toBe(false);
    change(rows[1]!.querySelector("input")!, "company.ghe.com");
    el("login-account-go").click();
    await vi.waitFor(() => expect(client.providerLogin).toHaveBeenCalledWith("POST", "opencode", expect.anything()));
    expect(client.providerLogin).toHaveBeenCalledWith("POST", "opencode", { choice: { accountId: "github-copilot", method: 0, inputs: { deploymentType: "enterprise", enterpriseUrl: "company.ghe.com" } } });
  });

  it("sends a typed key once, then clears the field", async () => {
    const { client } = open();
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    change(el("login-account-provider"), "openai");
    change(el("login-account-method"), "1");
    el("login-account-go").click();
    expect(client.providerLogin).not.toHaveBeenCalledWith("POST", expect.anything(), expect.anything());
    expect(el("login-status").textContent).toContain("API key");
    change(el("login-account-key"), "sk-secret");
    el("login-account-go").click();
    await vi.waitFor(() => expect(client.providerLogin).toHaveBeenCalledWith("POST", "opencode", { choice: { accountId: "openai", method: 1, key: "sk-secret" } }));
    await vi.waitFor(() => expect(el<HTMLInputElement>("login-account-key").value).toBe(""));
    expect(el("login-continue").hidden).toBe(false);
  });

  it("offers to sign out of an account that is signed in, after asking, and only that one", async () => {
    const { client } = open();
    el("panel-login").hidden = false;
    await vi.waitFor(() => expect(el("login-account-form").hidden).toBe(false));
    expect(el("login-account-logout").hidden).toBe(true);
    change(el("login-account-provider"), "github-copilot");
    expect(el("login-account-logout").hidden).toBe(false);
    el("login-account-logout").click();
    document.querySelector<HTMLButtonElement>("[data-confirm-accept]")!.click();
    await vi.waitFor(() => expect(client.providerLogout).toHaveBeenCalledWith("opencode", "github-copilot"));
    await vi.waitFor(() => expect(client.loginAccounts).toHaveBeenCalledTimes(2));
  });

  it("offers to try again when the accounts cannot be loaded", async () => {
    const { client } = open();
    client.loginAccounts.mockRejectedValueOnce(new Error("OpenCode did not start in time. Try again."));
    el("login-account-retry").click();
    await vi.waitFor(() => expect(el("login-account-loading").textContent).toContain("did not start"));
    expect(el("login-account-retry").hidden).toBe(false);
  });
});
