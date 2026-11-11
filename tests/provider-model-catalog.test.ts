import { expect, it, vi } from "vitest";
import { ProviderRegistry } from "../providers/src/registry.js";
import { MockAdapter } from "../providers/src/mock.js";
const catalog = (id: string) => ({ models: [{ id, name: id }], defaultModelId: id });
function fixture() {
 let revision = "one";
 class CatalogAdapter extends MockAdapter { runtimeRevision() { return revision; } async listModels() { return catalog("default"); } }
 const adapter = new CatalogAdapter({ id: "opencode" });
 const registry = new ProviderRegistry().register(adapter);
 const list = vi.spyOn(adapter, "listModels");
 return { registry, adapter, list, revision: (value: string) => { revision = value; } };
}
it("O01 shares one model-discovery process across simultaneous pane requests", async () => {
 const w = fixture(); let finish!: (value: ReturnType<typeof catalog>) => void;
 w.list.mockImplementation(() => new Promise(resolve => { finish = resolve; }));
 const a = w.registry.listModels("opencode"), b = w.registry.listModels("opencode"), c = w.registry.listModels("opencode");
 await Promise.resolve();
 expect(w.list).toHaveBeenCalledTimes(1); finish(catalog("one"));
 expect(await Promise.all([a,b,c])).toEqual([catalog("one"),catalog("one"),catalog("one")]);
 expect(await w.registry.listModels("opencode")).toEqual(catalog("one")); expect(w.list).toHaveBeenCalledTimes(1);
});
it("O01 does not cache an older discovery after a runtime switch", async () => {
 const w = fixture(); let old!: (value: ReturnType<typeof catalog>) => void;
 w.list.mockImplementationOnce(() => new Promise(resolve => { old = resolve; })).mockResolvedValue(catalog("new"));
 const previous = w.registry.listModels("opencode"); w.revision("two");
 expect(await w.registry.listModels("opencode")).toEqual(catalog("new")); old(catalog("old")); await previous;
 expect(await w.registry.listModels("opencode")).toEqual(catalog("new")); expect(w.list).toHaveBeenCalledTimes(2);
});
it("O01 allows retry after discovery fails and invalidates an in-flight catalog", async () => {
 const w = fixture(); w.list.mockRejectedValueOnce(Error("offline")).mockResolvedValue(catalog("ready"));
 await expect(w.registry.listModels("opencode")).rejects.toThrow("offline"); expect(await w.registry.listModels("opencode")).toEqual(catalog("ready"));
 let finish!: (value: ReturnType<typeof catalog>) => void; w.registry.invalidateModels("opencode");
 w.list.mockImplementationOnce(() => new Promise(resolve => { finish = resolve; })); const old = w.registry.listModels("opencode");
 w.registry.invalidateModels("opencode"); expect(await w.registry.listModels("opencode")).toEqual(catalog("ready"));
 finish(catalog("stale")); await old; expect(await w.registry.listModels("opencode")).toEqual(catalog("ready"));
});
