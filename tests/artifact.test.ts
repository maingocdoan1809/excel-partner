import { describe, expect, it } from "vitest";
import { CreateArtifactInputSchema, isExcelToolName } from "../protocol/src/index.js";
import { artifactKind, buildPreviewDocument, parseCsv } from "../addin/taskpane/ui/artifact.js";

describe("create_artifact input", () => {
  it("is a known tool", () => expect(isExcelToolName("create_artifact")).toBe(true));

  it.each(["report.html", "Notes 2.md", "data.csv", "a.SVG", "x.json", "t.txt"])("accepts %s", (filename) => {
    expect(CreateArtifactInputSchema.safeParse({ filename, content: "x" }).success).toBe(true);
  });

  it.each(["../x.html", "dir/x.html", "x.exe", "x", ".html", "x.html\\y", "x.js"])("rejects %s", (filename) => {
    expect(CreateArtifactInputSchema.safeParse({ filename, content: "x" }).success).toBe(false);
  });

  it("rejects empty and oversized content", () => {
    expect(CreateArtifactInputSchema.safeParse({ filename: "a.md", content: "" }).success).toBe(false);
    expect(CreateArtifactInputSchema.safeParse({ filename: "a.md", content: "x".repeat(300_001) }).success).toBe(false);
  });
});

describe("artifact preview document", () => {
  it("detects the kind from the extension", () => {
    expect(artifactKind("a.HTM")).toBe("html");
    expect(artifactKind("a.md")).toBe("markdown");
    expect(artifactKind("a.unknown")).toBe("text");
  });

  it("puts the network-closing policy ahead of the author's markup", () => {
    const doc = buildPreviewDocument({ name: "a.html", content: "<script>fetch('https://evil.example')</script>" });
    expect(doc.indexOf("Content-Security-Policy")).toBeLessThan(doc.indexOf("<script>"));
    expect(doc).toContain("default-src 'none'");
    expect(doc).not.toContain("connect-src");
  });

  it("keeps a leading doctype first so the page is not rendered in quirks mode", () => {
    const doc = buildPreviewDocument({ name: "a.html", content: "<!DOCTYPE html><html><body>hi</body></html>" });
    expect(doc.startsWith("<!DOCTYPE html>")).toBe(true);
    expect(doc.indexOf("Content-Security-Policy")).toBeGreaterThan(0);
  });

  it("escapes text and json instead of running them", () => {
    expect(buildPreviewDocument({ name: "a.txt", content: "<img src=x onerror=alert(1)>" })).not.toContain("<img");
    expect(buildPreviewDocument({ name: "a.json", content: '{"a":"<b>"}' })).toContain("&lt;b&gt;");
  });

  it("renders csv as a table with quoted commas and line breaks", () => {
    expect(parseCsv('a,b\n"x, y","line1\nline2"\r\n1,2')).toEqual([["a", "b"], ["x, y", "line1\nline2"], ["1", "2"]]);
    const doc = buildPreviewDocument({ name: "a.csv", content: "n,<v>\n1,2" });
    expect(doc).toContain("<th>n</th>");
    expect(doc).toContain("<th>&lt;v&gt;</th>");
  });
});
