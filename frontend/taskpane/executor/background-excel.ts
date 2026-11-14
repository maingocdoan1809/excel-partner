import { ToolFailure } from "./common.js";

/**
 * Keep generated Excel code from moving the user's viewport. This is a UX guard,
 * not a JavaScript sandbox: execute_office_js still has the pane's privileges.
 * Only objects handed to that call are wrapped; user navigation and other Office
 * batches keep their original methods, including while this call awaits sync().
 */
export function backgroundExcel() {
  const skipped = new Set<string>();
  const proxies = new WeakMap<object, object>();
  const originals = new WeakMap<object, object>();
  const navigation = new Set(["activate", "select", "setSelected", "setSelectedShapes", "suspendScreenUpdatingUntilNextSync"]);
  const implicitTargets = new Set(["getActiveWorksheet", "getSelectedRange", "getSelectedRanges", "getActiveCell", "getActiveChart", "getActiveChartOrNullObject", "getActiveShape"]);
  let navigationSkipped = false;
  const unwrap = (value: unknown): unknown => value !== null && typeof value === "object" ? originals.get(value) ?? value : value;

  function wrap<T>(value: T): T {
    if (value === null || typeof value !== "object" || value instanceof Promise || ArrayBuffer.isView(value)) return value;
    const cached = proxies.get(value);
    if (cached) return cached as T;
    // Collection items must also be wrapped when code uses map/forEach/iteration.
    if (Array.isArray(value)) return value.map(item => wrap(item)) as T;
    const proxy = new Proxy(value, {
      get(target, property) {
        const member = Reflect.get(target, property, target);
        if (typeof member !== "function") return wrap(member);
        if (navigation.has(String(property))) return () => {
          skipped.add(String(property));
          if (property !== "suspendScreenUpdatingUntilNextSync") navigationSkipped = true;
        };
        return (...args: unknown[]) => {
          // Selecting a target and then writing to the active selection would now
          // write to the USER's current cells. Refuse that pattern, never redirect it.
          if (navigationSkipped && implicitTargets.has(String(property))) {
            throw new ToolFailure("invalid_arguments", "This script tried to use the active sheet or selection after a suppressed activate/select call. Use explicit worksheet and range objects for every write; the user's viewport was not moved. Inspect any earlier changes before retrying.");
          }
          return wrap(Reflect.apply(member, target, args.map(unwrap)));
        };
      },
      set(target, property, next) { return Reflect.set(target, property, unwrap(next), target); },
    });
    proxies.set(value, proxy);
    originals.set(proxy, value);
    return proxy;
  }

  // The normal contract supplies context directly. Also guard a nested Excel.run
  // so a model using that familiar pattern does not unexpectedly move the view.
  const namespace = new Proxy(Excel, {
    get(target, property) {
      if (property !== "run") return Reflect.get(target, property, target);
      return (...args: unknown[]) => {
        const callback = args[args.length - 1];
        if (typeof callback === "function") args[args.length - 1] = (context: Excel.RequestContext) => callback(wrap(context));
        return Reflect.apply(target.run, target, args.map(unwrap));
      };
    },
  });
  return { wrap, namespace, skipped };
}
