/** Copilot can use the first prompt's injected context as its automatic name. */
export function cleanSessionTitle(provider: string, value: string | undefined): string | undefined {
  if (provider === "copilot" && value) {
    const request = value.split(/\s*<addin_context\b/i)[0]?.trim();
    if (!request || /<(?:selection_snapshot|saved_reference_files|user_attached_files|conversation_history)\b/i.test(request)) return undefined;
    return request;
  }
  return value;
}
