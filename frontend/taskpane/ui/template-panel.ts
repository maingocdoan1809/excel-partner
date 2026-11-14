import type { BridgeClient, SkillTemplate } from "../bridge-client.js";

export interface TemplatePanelDeps {
  getClient(): BridgeClient;
  elements: {
    templateList: HTMLElement;
    templateSkill: HTMLSelectElement;
    templateAdd: HTMLButtonElement;
  };
}

/** Example files attached to a user-installed skill. */
export function setupTemplatePanel(deps: TemplatePanelDeps) {
  const els = deps.elements;

  async function loadTemplateSkills(): Promise<void> {
    els.templateList.textContent = "Loading skills…";
    try {
      const selected = els.templateSkill.value;
      const skills = await deps.getClient().listTemplateSkills();
      els.templateSkill.replaceChildren(...skills.map((skill) => {
        const option = document.createElement("option");
        option.value = skill.name;
        option.textContent = skill.name;
        return option;
      }));
      if (skills.some((skill) => skill.name === selected)) els.templateSkill.value = selected;
      els.templateAdd.disabled = skills.length === 0;
      if (skills.length) await renderTemplates();
      else {
        const option = document.createElement("option");
        option.textContent = "No skills available";
        els.templateSkill.append(option);
        els.templateList.innerHTML = '<div class="empty"><strong>Add your first skill</strong><p>Open Skills in Settings to add a skill, then attach example files here.</p></div>';
      }
      els.templateSkill.disabled = skills.length === 0;
    } catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
  }

  async function renderTemplates(): Promise<void> {
    const skill = els.templateSkill.value;
    if (!skill) return;
    els.templateList.textContent = "Loading templates…";
    try {
      const templates = await deps.getClient().listSkillTemplates(skill);
      if (skill !== els.templateSkill.value) return;
      els.templateList.replaceChildren(...templates.map((item) => templateRow(skill, item)));
      if (!templates.length) els.templateList.textContent = "No templates for this skill yet.";
    } catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
  }

  function templateRow(skill: string, item: SkillTemplate): HTMLElement {
    const row = document.createElement("div");
    row.className = "template-row";
    const label = document.createElement("span");
    label.textContent = `${item.name} · ${item.kind} · ${Math.max(1, Math.round(item.size / 1024))} KB`;
    label.title = item.path;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "ghost";
    remove.textContent = "Delete";
    remove.setAttribute("aria-label", `Delete template ${item.name}`);
    remove.addEventListener("click", async () => {
      remove.disabled = true;
      try { await deps.getClient().deleteSkillTemplate(skill, item.name); await renderTemplates(); }
      catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
    });
    row.append(label, remove);
    return row;
  }

  async function addTemplate(file: File): Promise<void> {
    const skill = els.templateSkill.value;
    if (!skill) return;
    els.templateAdd.disabled = true;
    try { await deps.getClient().addSkillTemplate(skill, file); await renderTemplates(); }
    catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
    finally { els.templateAdd.disabled = false; }
  }

  return { loadTemplateSkills, renderTemplates, addTemplate };
}
