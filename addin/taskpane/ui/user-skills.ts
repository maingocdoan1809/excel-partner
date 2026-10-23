import type { BridgeClient } from "../bridge-client.js";
import { createTextEditor } from "./text-editor.js";

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
        copy.append(detail);
        row.append(copy);
        const toggle = document.createElement("button");
        toggle.type = "button";
        toggle.className = "secondary";
        toggle.textContent = skill.enabled ? "Disable" : "Enable";
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
        row.append(toggle);
        if (skill.removable) {
          const actions = document.createElement("div");
          actions.className = "user-skill-actions";
          const edit = document.createElement("button");
          edit.type = "button";
          edit.className = "secondary";
          edit.textContent = "Edit";
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
          const remove = document.createElement("button");
          remove.type = "button";
          remove.className = "secondary";
          remove.textContent = "Remove";
          remove.addEventListener("click", async () => {
            if (busy) return;
            if (remove.dataset.confirm !== "true") {
              remove.dataset.confirm = "true";
              remove.textContent = "Confirm";
              window.setTimeout(() => { remove.dataset.confirm = "false"; remove.textContent = "Remove"; }, 3000);
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
          row.append(actions);
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
