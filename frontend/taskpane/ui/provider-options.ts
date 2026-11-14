import type { ProviderOption } from "../../../protocol/src/provider-options.js";

/**
 * The settings the AI provider's own CLI offers for the open session (a mode such as plan or build, for example), drawn from what the
 * provider reports: a small heading and one row per choice, in the look of the other pickers. Nothing here knows what a setting means: a
 * choice goes back to the provider, which answers with the options as they now stand.
 */
export function setupProviderOptions(hooks: {
  container: HTMLElement;
  change(id: string, value: string): Promise<ProviderOption[]>;
  isBusy(): boolean;
  fail(message: string): void;
}) {
  let options: ProviderOption[] = [];
  let changing = false;

  const draw = () => {
    hooks.container.replaceChildren();
    hooks.container.hidden = options.length === 0;
    for (const option of options) {
      const group = document.createElement("div");
      group.className = "provider-option";
      group.setAttribute("role", "radiogroup");
      const title = document.createElement("div");
      title.className = "provider-option-title";
      title.textContent = option.label;
      group.setAttribute("aria-label", option.label);
      group.append(title);
      for (const choice of option.choices) {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "provider-choice";
        row.setAttribute("role", "radio");
        row.setAttribute("aria-checked", String(choice.value === option.value));
        row.dataset.optionId = option.id;
        row.dataset.value = choice.value;
        row.textContent = choice.label;
        row.disabled = hooks.isBusy() || changing;
        row.addEventListener("click", async () => {
          if (choice.value === option.value) return;
          changing = true;
          sync();
          try { options = await hooks.change(option.id, choice.value); }
          catch (error) { hooks.fail(error instanceof Error ? error.message : "Could not change the setting."); }
          finally { changing = false; draw(); }
        });
        group.append(row);
      }
      hooks.container.append(group);
    }
  };
  /** The rows cannot be used while a reply is running or a change is on its way. */
  const sync = () => {
    for (const row of hooks.container.querySelectorAll<HTMLButtonElement>(".provider-choice")) row.disabled = hooks.isBusy() || changing;
  };

  draw();
  return {
    /** The provider's current options (an empty list hides the section). */
    render(next: ProviderOption[]): void { options = next; draw(); },
    /** A new chat or another provider: what the last session offered no longer applies. */
    clear(): void { options = []; draw(); },
    sync,
  };
}
