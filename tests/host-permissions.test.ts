import { afterEach, describe, expect, it, vi } from "vitest";
import { executeWithPermission } from "../addin/taskpane/permissions.js";
import { executeWordTool } from "../addin/taskpane/executor/word/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";
import { WORD_WRITE_TOOLS, PPT_WRITE_TOOLS } from "../protocol/src/index.js";

afterEach(() => uninstallWordFake());
describe("Ask permissions across Office hosts", () => {
  it("does not overwrite Word text when approval is declined", async () => {
    const doc = installWordFake(new FakeWordDoc(["Keep this paragraph."]));
    const ask = vi.fn(async () => false);
    const result = await executeWithPermission("edit_doc_text", { old_text: "Keep this paragraph.", new_text: "Replaced." }, "ask", executeWordTool, ask);
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(ask).toHaveBeenCalledOnce();
    expect(doc.texts()).toEqual(["Keep this paragraph."]);
  });
  it("applies Word text after approval and automatically in Auto", async () => {
    const doc = installWordFake(new FakeWordDoc(["Original"]));
    expect((await executeWithPermission("edit_doc_text", { old_text: "Original", new_text: "Approved" }, "ask", executeWordTool, async () => true)).ok).toBe(true);
    const ask = vi.fn(async () => false);
    expect((await executeWithPermission("edit_doc_text", { old_text: "Approved", new_text: "Automatic" }, "auto", executeWordTool, ask)).ok).toBe(true);
    expect(doc.texts()).toEqual(["Automatic"]);
    expect(ask).not.toHaveBeenCalled();
  });
  it.each([...WORD_WRITE_TOOLS, ...PPT_WRITE_TOOLS, "execute_office_js"].filter(tool => tool !== "propose_doc_edits"))(
    "does not execute %s before approval", async tool => {
      const execute = vi.fn(async () => ({ ok: true as const, data: {} }));
      const result = await executeWithPermission(tool, {}, "ask", execute, async () => false);
      expect(result.ok).toBe(false);
      expect(execute).not.toHaveBeenCalled();
    });
  it("leaves read tools and the proposal's own review UI in charge", async () => {
    const ask = vi.fn(async () => false);
    const execute = vi.fn(async () => ({ ok: true as const, data: {} }));
    for (const tool of ["read_doc_section", "read_slide_text", "propose_doc_edits"]) {
      expect((await executeWithPermission(tool, {}, "ask", execute, ask)).ok).toBe(true);
    }
    expect(ask).not.toHaveBeenCalled();
  });
});
