import type { BridgeClient } from "../bridge-client.js";
import { createTextEditor } from "./text-editor.js";

/** A small square button that shows a symbol; the label lives in its tooltip so the row stays compact. */
function iconButton(symbol: string, label: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "secondary skill-icon";
  button.textContent = symbol;
  button.title = label;
  return button;
}

export function setupUserSkills(client: () => BridgeClient, changed: () => void): { refresh(): Promise<void> } {
  const list = document.getElementById("user-skills-list")!;
  const message = document.getElementById("user-skills-status")!;
  const upload = document.getElementById("skill-upload-button") as HTMLButtonElement;
  const input = document.getElementById("skill-upload-input") as HTMLInputElement;
  let busy = false;
  let request = 0;
  const editor = createTextEditor();

  async function refresh(): Promise<void> {
    const current = ++request;
    list.textContent = "Loading skills…";
    try {
      const skills = await client().listSkills();
      if (current !== request) return;
      list.replaceChildren();
      if (!skills.length) { list.textContent = "No skills installed. Upload your own to get started."; return; }
      for (const skill of skills) {
        const row = document.createElement("div");
        row.className = "user-skill";
        row.dataset.enabled = String(skill.enabled);
        const head = document.createElement("div");
        head.className = "user-skill-head";
        const copy = document.createElement("div");
        copy.className = "user-skill-copy";
        const title = document.createElement("strong");
        title.textContent = `/${skill.name}`;
        const detail = document.createElement("p");
        detail.textContent = skill.description || "Custom skill";
        copy.append(title);
        if (!skill.enabled) {
          const badge = document.createElement("span");
          badge.className = "skill-disabled-label";
          badge.textContent = "Disabled";
          copy.append(badge);
        }
        const actions = document.createElement("div");
        actions.className = "user-skill-actions";
        head.append(copy, actions);
        row.append(head, detail);
        const toggle = iconButton(skill.enabled ? "⏻" : "○", skill.enabled ? "Disable this skill" : "Enable this skill");
        toggle.dataset.on = String(skill.enabled);
        toggle.setAttribute("role", "switch");
        toggle.setAttribute("aria-checked", String(skill.enabled));
        toggle.setAttribute("aria-label", `Enable /${skill.name}`);
        toggle.addEventListener("click", async () => {
          if (busy) return;
          busy = true;
          toggle.disabled = true;
          try {
            await client().setSkillEnabled(skill.name, !skill.enabled);
            message.textContent = `/${skill.name} ${skill.enabled ? "disabled" : "enabled"}. Previously loaded instructions remain in chat history.`;
            changed();
            await refresh();
          } catch (error) { message.textContent = error instanceof Error ? error.message : "Could not update skill."; }
          finally { busy = false; toggle.disabled = false; }
        });
        actions.append(toggle);
        if (skill.removable) {
          const edit = iconButton("✎", "Edit this skill");
          edit.setAttribute("aria-label", `Edit /${skill.name}`);
          edit.addEventListener("click", async () => {
            if (busy) return;
            busy = true;
            edit.disabled = true;
            try {
              const files = await client().skillTextFiles(skill.name);
              await editor.open({ title: `Edit /${skill.name}`, detail: "Edit text files in this skill. Workbook templates and images are preserved.", files,
                load: file => client().readSkillText(skill.name, file),
                save: document => client().saveSkillText(skill.name, document),
                saved: () => { changed(); void refresh(); },
              });
            } catch (error) { message.textContent = error instanceof Error ? error.message : "Could not edit this skill."; }
            finally { busy = false; edit.disabled = false; }
          });
          const remove = iconButton("🗑", "Remove this skill");
          remove.setAttribute("aria-label", `Remove /${skill.name}`);
          remove.addEventListener("click", async () => {
            if (busy) return;
            if (remove.dataset.confirm !== "true") {
              remove.dataset.confirm = "true";
              remove.textContent = "✓";
              remove.title = "Click again to remove";
              window.setTimeout(() => { remove.dataset.confirm = "false"; remove.textContent = "🗑"; remove.title = "Remove this skill"; }, 3000);
              return;
            }
            busy = true;
            upload.disabled = true;
            remove.disabled = true;
            try {
              await client().removeSkill(skill.name);
              message.textContent = `Removed /${skill.name}.`;
              changed();
              await refresh();
            } catch (error) { message.textContent = error instanceof Error ? error.message : "Could not remove this skill."; }
            finally { busy = false; upload.disabled = false; remove.disabled = false; }
          });
          actions.append(edit, remove);
        }
        list.append(row);
      }
    } catch (error) { if (current === request) list.textContent = error instanceof Error ? error.message : "Could not load skills."; }
  }

  upload.addEventListener("click", () => { if (!busy) input.click(); });
  input.addEventListener("change", async () => {
    const file = input.files?.[0];
    input.value = "";
    if (!file || busy) return;
    busy = true;
    upload.disabled = true;
    message.textContent = "Uploading skill…";
    try {
      await client().uploadSkill(file);
      message.textContent = "Skill uploaded. Type / in chat to use it.";
      changed();
      await refresh();
    } catch (error) { message.textContent = error instanceof Error ? error.message : "Could not upload this skill."; }
    finally { busy = false; upload.disabled = false; }
  });
  return { refresh };
}
