/** A provider may take the first prompt, with the context the bridge injected into it, as the chat's automatic name. */
export function cleanSessionTitle(value: string | undefined): string | undefined {
  if (value) {
    const request = value.split(/\s*<addin_context\b/i)[0]?.trim();
    if (!request || /<(?:selection_snapshot|saved_reference_files|user_attached_files|conversation_history)\b/i.test(request)) return undefined;
    return request;
  }
  return value;
}
