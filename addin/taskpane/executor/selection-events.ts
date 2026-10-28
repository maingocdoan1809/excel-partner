import type { SelectionLabel } from "../ui/selection.js";

type Subscription = { context: Excel.RequestContext; remove(): void };

async function removeSubscriptions(subscriptions: Subscription[]): Promise<void> {
  const contexts = new Set(subscriptions.map(item => item.context));
  for (const context of contexts) {
    try {
      await Excel.run(context, async ctx => {
        for (const subscription of [...subscriptions].reverse()) if (subscription.context === context) subscription.remove();
        await ctx.sync();
      });
    } catch {
      // A deleted shape/sheet may no longer accept removal of its event handler.
    }
  }
}

/** Native shape events are registered per object, including children of groups.
 * Refresh subscriptions on worksheet activation, workbook edits and pane focus;
 * there is no timer scanning the workbook. */
export function watchShapeSelectionEvents(options: { onSelection(): void; onActivated?(value: SelectionLabel): void; onError(error: unknown): void }) {
  const shapes = new Map<string, Subscription[]>();
  const labels = new Map<string, SelectionLabel>();
  let activeKey: string | null = null;
  const workbookSubscriptions: Subscription[] = [];
  let stopped = false;
  let pending: Promise<void> | null = null;
  let refreshAgain = false;

  async function scan(): Promise<void> {
    const cleanup: Subscription[] = [];
    try { await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("id");
      sheet.shapes.load("items/id,items/name,items/type");
      await context.sync();
      const current = new Map<string, Excel.Shape>();
      async function visit(items: Excel.Shape[], depth = 0): Promise<void> {
        if (depth > 32) throw new Error("Shape groups exceed the supported nesting depth (32).");
        for (const shape of items) {
          if ((shape.type as string) === "Unsupported") continue; // charts: no native shape handlers
          const key = `${sheet.id}/${shape.id}`;
          current.set(key, shape);
          labels.set(key, { label: shape.name, title: `Selected ${shape.type === "Line" ? "line / connector" : "shape"}: ${shape.name} (${shape.id})` });
          if (shape.type === "Group") {
            shape.group.shapes.load("items/id,items/name,items/type");
            await context.sync();
            await visit(shape.group.shapes.items, depth + 1);
          }
        }
      }
      await visit(sheet.shapes.items);
      if (stopped) return;
      const additions = new Map<string, Subscription[]>();
      try {
        for (const [key, shape] of current) {
          if (shapes.has(key)) continue;
          const subscriptions: Subscription[] = [];
          additions.set(key, subscriptions);
          // Native handlers outlive this Excel.run batch. Keep their source object
          // alive until both event registrations have been removed.
          context.trackedObjects.add(shape);
          subscriptions.push({ context, remove: () => context.trackedObjects.remove(shape) });
          const activated = shape.onActivated.add(async () => {
            if (stopped) return;
            activeKey = key;
            const label = labels.get(key);
            if (label && options.onActivated) options.onActivated(label);
            else options.onSelection();
          });
          subscriptions.push({ context, remove: () => activated.remove() });
          const deactivated = shape.onDeactivated.add(async () => {
            if (stopped || (activeKey !== null && activeKey !== key)) return;
            activeKey = null;
            options.onSelection();
            // Do not rescan or mutate subscriptions during Excel's selection transition.
          });
          subscriptions.push({ context, remove: () => deactivated.remove() });
        }
        await context.sync();
      } catch (error) {
        cleanup.push(...[...additions.values()].flat());
        throw error;
      }
      for (const [key, subscriptions] of additions) shapes.set(key, subscriptions);
      for (const [key, subscriptions] of shapes) {
        if (!current.has(key)) { cleanup.push(...subscriptions); shapes.delete(key); labels.delete(key); }
      }
    }); } finally { await removeSubscriptions(cleanup); }
  }

  function refresh(): Promise<void> {
    if (stopped || !Office.context.requirements.isSetSupported("ExcelApi", "1.9")) return Promise.resolve();
    if (pending) { refreshAgain = true; return pending; }
    pending = (async () => {
      do {
        refreshAgain = false;
        try { await scan(); } catch (error) { if (!stopped) options.onError(error); }
      } while (refreshAgain && !stopped);
    })().finally(() => { pending = null; });
    return pending;
  }

  const ready = (async () => {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) return;
    try {
      await Excel.run(async context => {
        const activated = context.workbook.worksheets.onActivated.add(async () => {
          if (stopped) return;
          activeKey = null;
          options.onSelection();
          void refresh();
        });
        workbookSubscriptions.push({ context, remove: () => activated.remove() });
        await context.sync();
      });
      if (!stopped) await refresh();
    } catch (error) { if (!stopped) options.onError(error); }
  })();

  return {
    ready,
    refresh,
    async stop() {
      stopped = true;
      await ready;
      await pending;
      const subscriptions = [...workbookSubscriptions, ...[...shapes.values()].flat()];
      shapes.clear();
      labels.clear();
      workbookSubscriptions.length = 0;
      await removeSubscriptions(subscriptions);
    },
  };
}
