import type { AdapterClass } from "../../providers/src/profiled.js";
import { MockAdapter } from "../../providers/src/mock.js";
import type { ProviderAdapter, ProviderRuntime } from "../../providers/src/types.js";

/** Wraps a function that builds an adapter as the class a ProfiledAdapter expects. */
export function adapterClass(build: (runtime: ProviderRuntime) => ProviderAdapter): AdapterClass {
  return class { static readonly capabilities = MockAdapter.capabilities; constructor(runtime: ProviderRuntime) { return build(runtime) as never; } } as unknown as AdapterClass;
}
