import fs from "node:fs";
import { describe, expect, it } from "vitest";

describe("task pane content security policy", () => {
  const html = fs.readFileSync("frontend/taskpane/index.html", "utf8");
  const policy = /http-equiv="Content-Security-Policy" content="([^"]+)"/.exec(html)?.[1] ?? "";
  const directive = (name: string) => policy.split(";").map(part => part.trim()).find(part => part.startsWith(`${name} `)) ?? "";

  it("lets the pane connect only to the bridge that serves it", () => {
    expect(directive("connect-src")).toBe("connect-src 'self'");
    expect(directive("default-src")).toBe("default-src 'self'");
    expect(directive("object-src")).toBe("object-src 'none'");
  });

  it("allows Office.js and nothing else from outside for scripts besides the preview CDNs", () => {
    expect(html.indexOf("Content-Security-Policy")).toBeLessThan(html.indexOf("<script"));
    expect(directive("script-src").split(" ").filter(source => source.startsWith("https:")).sort())
      .toEqual(["https://appsforoffice.microsoft.com", "https://cdn.jsdelivr.net", "https://cdnjs.cloudflare.com"]);
  });
});
