import { CANCELLED_USER_INPUT, type UserInputRequest, type UserInputResponse } from "../../../protocol/src/index.js";

/** A native question lives only for its provider request, never across reconnects. */
export function showUserInput(container: HTMLElement, request: UserInputRequest, signal: AbortSignal): Promise<UserInputResponse> {
  if (signal.aborted) return Promise.resolve(CANCELLED_USER_INPUT);
  return new Promise(resolve => {
    const form = document.createElement("form");
    form.className = "user-input-card";
    form.setAttribute("aria-label", "Assistant questions");
    const fields = request.questions.map((question, index) => {
      const fieldset = document.createElement("fieldset");
      const legend = document.createElement("legend");
      legend.textContent = question.header || `Question ${index + 1}`;
      const text = document.createElement("p");
      text.textContent = question.question;
      fieldset.append(legend, text);
      const inputs = question.options.map(option => {
        const label = document.createElement("label");
        label.className = "user-input-option";
        const input = document.createElement("input");
        input.type = question.multiSelect ? "checkbox" : "radio";
        input.name = `question-${index}`;
        input.value = option.label;
        const copy = document.createElement("span");
        const title = document.createElement("strong");
        title.textContent = option.label;
        copy.append(title);
        if (option.description) {
          const description = document.createElement("small");
          description.textContent = option.description;
          copy.append(description);
        }
        label.append(input, copy);
        fieldset.append(label);
        return input;
      });
      let freeform: HTMLInputElement | undefined;
      if (question.allowFreeform || inputs.length === 0) {
        const label = document.createElement("label");
        label.className = "user-input-freeform";
        label.append(inputs.length ? "Other answer" : "Your answer");
        freeform = document.createElement("input");
        freeform.type = question.isSecret ? "password" : "text";
        freeform.maxLength = 20_000;
        freeform.autocomplete = "off";
        label.append(freeform);
        fieldset.append(label);
        freeform.addEventListener("input", () => {
          if (!question.multiSelect && freeform!.value.trim()) for (const input of inputs) input.checked = false;
        });
        for (const input of inputs) input.addEventListener("change", () => {
          if (!question.multiSelect && input.checked) freeform!.value = "";
        });
      }
      form.append(fieldset);
      return { question, inputs, freeform };
    });
    const error = document.createElement("p");
    error.className = "user-input-error";
    error.setAttribute("role", "alert");
    const actions = document.createElement("div");
    actions.className = "user-input-actions";
    const skip = document.createElement("button");
    skip.type = "button";
    skip.textContent = "Skip";
    const submit = document.createElement("button");
    submit.type = "submit";
    submit.textContent = "Reply";
    actions.append(skip, submit);
    form.append(error, actions);
    let finished = false;
    const finish = (response: UserInputResponse) => {
      if (finished) return;
      finished = true;
      signal.removeEventListener("abort", cancel);
      form.remove();
      resolve(response);
    };
    const cancel = () => finish(CANCELLED_USER_INPUT);
    skip.addEventListener("click", cancel);
    form.addEventListener("submit", event => {
      event.preventDefault();
      const answers = Object.create(null) as Record<string, string[]>;
      for (const { question, inputs, freeform } of fields) {
        const selected = inputs.filter(input => input.checked).map(input => input.value);
        const typed = freeform?.value.trim();
        if (typed) selected.push(typed);
        if (!selected.length) { error.textContent = "Please answer every question, or choose Skip."; return; }
        answers[question.id] = [...new Set(selected)];
      }
      finish({ cancelled: false, answers });
    });
    signal.addEventListener("abort", cancel, { once: true });
    container.append(form);
    form.querySelector<HTMLInputElement>("input")?.focus();
  });
}
