/** Bound the time between checkpoints even while streaming never pauses. */
export function createHistoryScheduler(persist: () => void, interval = 500) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return {
    schedule() { if (timer === undefined) timer = setTimeout(() => { timer = undefined; persist(); }, interval); },
    flush() { if (timer !== undefined) clearTimeout(timer); timer = undefined; persist(); },
  };
}
