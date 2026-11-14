export interface SelectionLabel {
  label: string;
  title: string;
}

/** Serialize event-driven reads and discard results invalidated by a newer event. */
export function monitorSelection(options: {
  read(): Promise<SelectionLabel>;
  render(value: SelectionLabel): void;
  onError(error: unknown): void;
  isVisible(): boolean;
}) {
  let running = false;
  let stopped = false;
  let requested = false;
  let revision = 0;
  let previous: SelectionLabel | undefined;

  function show(value: SelectionLabel): void {
    if (stopped || !options.isVisible()) return;
    if (previous?.label !== value.label || previous.title !== value.title) {
      options.render(value);
      previous = value;
    }
  }

  async function refresh(): Promise<void> {
    if (stopped || running || !options.isVisible()) return;
    running = true;
    const readingRevision = revision;
    try {
      const value = await options.read();
      if (!stopped && readingRevision === revision) show(value);
    } catch (error) {
      if (!stopped && readingRevision === revision) {
        previous = undefined;
        options.onError(error);
      }
    } finally {
      running = false;
      if (requested && !stopped) { requested = false; void refresh(); }
    }
  }

  void refresh();
  return {
    show(value: SelectionLabel) {
      revision++;
      requested = false;
      show(value);
    },
    refresh() {
      revision++;
      if (running) requested = true;
      else void refresh();
    },
    stop() { stopped = true; },
  };
}
