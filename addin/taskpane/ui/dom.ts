export function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`missing element #${id}`);
  return element as T;
}

/** Message of a caught error, or `fallback` when it is not an Error. */
export function errorText(error: unknown, fallback: string): string {
  return error instanceof Error ? error.message : fallback;
}
