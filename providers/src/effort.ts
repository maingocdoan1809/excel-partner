export interface EffortConfigOption {
  id: string;
  type: string;
  category?: string;
  currentValue?: string | boolean;
  options?: Array<{ value?: string; options?: Array<{ value: string }> }>;
}

export function effortConfig(options: EffortConfigOption[] | undefined, requested: string): { configId: string; value: string } | null {
  const option = options?.find(item => item.type === "select" && (item.category === "thought_level" || /effort|reasoning/i.test(item.id)));
  if (!option) return null;
  const values = (option.options ?? []).flatMap(item => item.options?.map(entry => entry.value) ?? (item.value ? [item.value] : []));
  const value = values.includes(requested) ? requested : values.includes("medium") ? "medium" : null;
  return value && value !== option.currentValue ? { configId: option.id, value } : null;
}
