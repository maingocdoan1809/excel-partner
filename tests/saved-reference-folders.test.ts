import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ToolResult } from "../protocol/src/index.js";
import { SavedReferences, type SavedReferenceStatus } from "../bridge/src/saved-references.js";
import { ReferenceFiles } from "../bridge/src/reference-files.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

const made: string[] = [];
let bridge: Bridge | undefined;
const panes: PaneClient[] = [];
afterEach(async () => {
  await Promise.all(panes.splice(0).map((pane) => pane.close()));
  await bridge?.close();
  bridge = undefined;
  for (const folder of made.splice(0)) fs.rmSync(folder, { recursive: true, force: true });
});

/** A folder with a file at its top, a file in a subfolder, and a sibling folder that must stay closed. */
function tree() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "saved-reference-folder-"));
  made.push(root);
  const docs = path.join(root, "docs");
  fs.mkdirSync(path.join(docs, "deep"), { recursive: true });
  fs.mkdirSync(path.join(root, "private"));
  fs.writeFileSync(path.join(docs, "rules.md"), "Top rules");
  fs.writeFileSync(path.join(docs, "deep", "more.md"), "Deep rules");
  fs.writeFileSync(path.join(root, "private", "secret.md"), "Secret");
  fs.writeFileSync(path.join(root, "outside.md"), "Outside");
  return { root, docs, metadata: path.join(root, "settings", "references.json") };
}

function accessFor(store: SavedReferences): ReferenceFiles {
  const access = new ReferenceFiles([]);
  access.setSavedPaths(() => store.paths());
  return access;
}

describe("saved reference folders", () => {
  it("opens everything inside a saved folder, and nothing beside or above it", () => {
    const { root, docs } = tree();
    const store = new SavedReferences(null);
    const entry = store.save({ name: "Docs", path: docs, description: "House rules", kind: "folder" });
    expect(entry.kind).toBe("folder");
    const access = accessFor(store);
    expect(access.resolve(docs)).toBe(fs.realpathSync(docs));
    expect(access.readSelectedFile(path.join(docs, "rules.md")).bytes.toString()).toBe("Top rules");
    expect(access.readSelectedFile(path.join(docs, "deep", "more.md")).bytes.toString()).toBe("Deep rules");
    expect(access.list(docs).entries.map((e) => `${e.type}:${e.name}`)).toEqual(["folder:deep", "file:rules.md"]);
    for (const closed of [path.join(root, "private", "secret.md"), path.join(root, "outside.md"), root]) {
      expect(() => access.resolve(closed), closed).toThrow("outside");
    }
    // A folder whose name only starts like the saved one is not inside it.
    fs.mkdirSync(`${docs}-extra`);
    fs.writeFileSync(path.join(`${docs}-extra`, "x.md"), "x");
    expect(() => access.resolve(path.join(`${docs}-extra`, "x.md"))).toThrow("outside");
    store.remove(entry.id);
    expect(() => access.resolve(path.join(docs, "rules.md"))).toThrow("outside");
  });

  it("does not let a link inside a saved folder lead out of it", () => {
    const { root, docs } = tree();
    const store = new SavedReferences(null);
    store.save({ name: "Docs", path: docs, description: "", kind: "folder" });
    const access = accessFor(store);
    // A junction needs no privilege on Windows; elsewhere it is an ordinary symlink.
    fs.symlinkSync(path.join(root, "private"), path.join(docs, "link"), "junction");
    expect(() => access.resolve(path.join(docs, "link", "secret.md"))).toThrow("outside");
  });

  it("opens nothing when the saved folder itself has been replaced by a link elsewhere", () => {
    const { root, docs } = tree();
    const store = new SavedReferences(null);
    store.save({ name: "Docs", path: docs, description: "", kind: "folder" });
    const access = accessFor(store);
    fs.rmSync(docs, { recursive: true });
    fs.symlinkSync(path.join(root, "private"), docs, "junction");
    expect(() => access.resolve(path.join(docs, "secret.md"))).toThrow("outside");
    expect(store.list()[0]!.available).toBe(false);
  });

  it("works for a folder at the top of a drive, whose path already ends with a separator", () => {
    const top = path.parse(process.cwd()).root;
    const store = new SavedReferences(null);
    store.save({ name: "Everything", path: top, description: "", kind: "folder" });
    expect(accessFor(store).resolve(os.tmpdir())).toBe(fs.realpathSync(os.tmpdir()));
  });

  it("tells a file from a folder in the list, marks a missing one unavailable and keeps the kind through an edit", () => {
    const { docs, metadata } = tree();
    const file = path.join(docs, "rules.md");
    const store = new SavedReferences(metadata);
    const folder = store.save({ name: "Docs", path: docs, description: "", kind: "folder" });
    store.save({ name: "Rules", path: file, description: "", kind: "file" });
    expect(new SavedReferences(metadata).list().map((e) => [e.name, e.kind, e.available])).toEqual([["Docs", "folder", true], ["Rules", "file", true]]);
    // An edit that gives no new kind keeps what it was.
    store.save({ id: folder.id, name: "Docs renamed", path: docs, description: "x" });
    expect(store.list()[0]).toMatchObject({ name: "Docs renamed", kind: "folder" });
    fs.rmSync(docs, { recursive: true });
    expect(store.list().map((e) => [e.kind, e.available])).toEqual([["folder", false], ["file", false]]);
    // A folder saved as a file is not available either.
    const wrong = new SavedReferences(null);
    wrong.save({ name: "Wrong", path: os.tmpdir(), description: "", kind: "file" });
    expect(wrong.list()[0]!.available).toBe(false);
  });

  it("still reads references saved before folders existed, as files, and refuses a damaged kind", () => {
    const { docs, metadata } = tree();
    const file = path.join(docs, "rules.md");
    fs.mkdirSync(path.dirname(metadata), { recursive: true });
    fs.writeFileSync(metadata, JSON.stringify([{ id: "old", name: "Old", path: file, description: "from before" }]));
    expect(new SavedReferences(metadata).list()).toEqual([{ id: "old", name: "Old", path: file, description: "from before", kind: "file", available: true }]);
    fs.writeFileSync(metadata, JSON.stringify([{ id: "x", name: "X", path: file, description: "", kind: "drive" }]));
    expect(() => new SavedReferences(metadata)).toThrow("Invalid saved reference metadata");
  });

  it("tells the assistant which are folders and how to read inside them", () => {
    const { docs } = tree();
    const store = new SavedReferences(null);
    store.save({ name: "Docs", path: docs, description: "House rules", kind: "folder" });
    const prompt = store.prompt();
    expect(prompt).toContain("<saved_reference_files>");
    expect(prompt).toContain('"kind":"folder"');
    expect(prompt).toContain("list_reference_files");
    expect(prompt).toContain("nothing outside it");
  });

  it("takes a folder typed in the pane, and reads inside it through the tools without asking", async () => {
    const { docs, root, metadata } = tree();
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    bridge = new Bridge(config, { savedReferencesFile: metadata, references: new ReferenceFiles([]), usageFile: null,
      registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })) });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    panes.push(pane);
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "reference-folder-123" });
    const hello = await pane.waitFor((m) => m.type === "hello_ok");
    if (hello.type !== "hello_ok") throw new Error("No hello");
    const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
    const request = (route: string, method = "GET", body?: unknown) => fetch(`${url}${route}`, { method, headers, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
    const tool = async (name: string, args: unknown) => (await (await request("/internal/tool", "POST", { tool: name, args })).json()) as ToolResult;

    expect((await request("/session/references", "POST", { name: "Docs", description: "Rules", path: path.join(root, "missing") })).status).toBe(400);
    expect((await request("/session/references", "POST", { name: "Docs", description: "Rules", path: `  "${docs}"  ` })).status).toBe(200);
    const listed = await (await request("/session/references")).json() as { references: SavedReferenceStatus[] };
    expect(listed.references).toHaveLength(1);
    expect(listed.references[0]).toMatchObject({ name: "Docs", kind: "folder", available: true });
    expect((await request("/session/references", "POST", { name: "Docs again", description: "", path: docs })).status).toBe(400);

    // The pane answers an access question with a refusal, so any question that is asked shows as a refusal.
    let questions = 0;
    pane.onMessage = (message) => {
      if (message.type === "tool_request" && message.tool === "__authorize_reference_path") {
        questions++;
        pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { approved: false } } });
      }
    };
    const inside = await tool("list_reference_files", { path: docs });
    expect(inside.ok).toBe(true);
    expect(JSON.stringify(inside)).toContain("rules.md");
    expect(JSON.stringify(await tool("read_attached_file", { path: path.join(docs, "deep", "more.md") }))).toContain("Deep rules");
    expect(questions).toBe(0);
    // Outside it the usual question is asked, and a refusal stays a refusal.
    expect((await tool("read_attached_file", { path: path.join(root, "private", "secret.md") })).ok).toBe(false);
    expect(questions).toBe(1);

    const id = listed.references[0]!.id;
    expect((await request("/session/references", "POST", { id, name: "Docs renamed", description: "Edited" })).status).toBe(200);
    expect(((await (await request("/session/references")).json()) as { references: SavedReferenceStatus[] }).references[0]).toMatchObject({ name: "Docs renamed", kind: "folder" });
    expect((await request("/session/references", "DELETE", { id })).status).toBe(200);
    expect((await tool("read_attached_file", { path: path.join(docs, "rules.md") })).ok).toBe(false);
  });
});
