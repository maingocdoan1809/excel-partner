import type { ProviderOption } from "../../protocol/src/provider-options.js";

/** A config option as an ACP agent describes it in `session/new` and `session/set_config_option`. */
export interface AcpConfigOption {
  id: string;
  name?: string;
  type: string;
  category?: string;
  currentValue?: string | boolean;
  options?: Array<{ value?: string; name?: string; options?: Array<{ value: string; name?: string }> }>;
}

/**
 * The options of the given categories that offer a choice, in the shape the pane draws. A provider names the categories it passes on:
 * the model and the effort already have their own controls, and permissions for Office edits are the add-in's.
 */
export function acpOptions(options: AcpConfigOption[] | undefined, categories: readonly string[]): ProviderOption[] {
  const result: ProviderOption[] = [];
  for (const option of options ?? []) {
    if (option.type !== "select" || !option.category || !categories.includes(option.category)) continue;
    const choices = (option.options ?? []).flatMap(item => item.options ?? (item.value === undefined ? [] : [{ value: item.value, name: item.name }]))
      .map(entry => ({ value: entry.value, label: entry.name?.trim() || entry.value }));
    if (!choices.length) continue;
    result.push({ id: option.id, label: option.name?.trim() || option.id, value: String(option.currentValue ?? ""), choices });
  }
  return result;
}
