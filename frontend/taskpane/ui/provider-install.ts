import type { ProviderId } from "../../../protocol/src/index.js";
import { byId, errorText } from "./dom.js";

export interface ProviderInstallOptions {
  client: () => { systemAction(action: string, provider?: ProviderId): Promise<void> };
  provider: () => ProviderId;
  isBusy: () => boolean;
  /** The private CLI is on this computer. `needsSignIn` when it is installed but still has to be signed in. */
  onInstalled: (provider: ProviderId, needsSignIn: boolean) => void;
}

/**
 * The Install button of the set-up panel: the same action as "Install private CLI" in Settings > System & recovery,
 * so a provider that is not installed can be set up without a terminal.
 */
export function setupProviderInstall(options: ProviderInstallOptions): { reset(): void } {
  const button = byId<HTMLButtonElement>("login-install");
  const status = byId("login-install-status");
  let working = false;

  function say(text: string, state?: "error"): void {
    status.textContent = text;
    if (state) status.dataset.state = state; else delete status.dataset.state;
  }

  button.addEventListener("click", async () => {
    if (working) return;
    if (options.isBusy()) { say("Wait for the current response to finish.", "error"); return; }
    const target = options.provider();
    working = true; button.disabled = true;
    say("Installing and verifying the private CLI. This may take a few minutes…");
    let needsSignIn = false;
    try {
      await options.client().systemAction("install", target);
    } catch (error) {
      // A private CLI that is installed but not signed in is reported as an error; it is installed all the same.
      if ((error as { code?: string }).code === "login_required") needsSignIn = true;
      else { say(errorText(error, "The installation did not finish. Try again."), "error"); working = false; button.disabled = false; return; }
    }
    working = false; button.disabled = false;
    say("");
    options.onInstalled(target, needsSignIn);
  });

  return { reset: () => { if (!working) say(""); } };
}
