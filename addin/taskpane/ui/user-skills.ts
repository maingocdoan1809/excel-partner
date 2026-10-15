import type { BridgeClient } from "../bridge-client.js";

export function setupUserSkills(client: () => BridgeClient, changed: () => void): { refresh(): Promise<void> } {
  const list = document.getElementById("user-skills-list")!;
  const message = document.getElementById("user-skills-status")!;
  const upload = document.getElementById("skill-upload-button") as HTMLButtonElement;
  const input = document.getElementById("skill-upload-input") as HTMLInputElement;
  let busy = false;
  let request = 0;

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
        const copy = document.createElement("div");
        const title = document.createElement("strong");
        title.textContent = `/${skill.name}`;
        const detail = document.createElement("p");
        detail.textContent = skill.description || "Custom skill";
        copy.append(title, detail);
        row.append(copy);
        if (skill.removable) {
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
          row.append(remove);
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
