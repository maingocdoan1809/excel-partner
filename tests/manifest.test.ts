// @vitest-environment jsdom
import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

const xml = fs.readFileSync(path.join(__dirname, "..", "addin", "manifest.xml"), "utf8");
const doc = new DOMParser().parseFromString(xml, "application/xml");
const all = (root: ParentNode, selector: string) => [...root.querySelectorAll(selector)];

describe("add-in manifest", () => {
  it("parses", () => {
    expect(doc.querySelector("parsererror")).toBeNull();
  });

  it("keeps the add-in id so an existing install is replaced, not duplicated, and has a newer version", () => {
    expect(doc.querySelector("OfficeApp > Id")?.textContent).toBe("7f2b6c1e-9a4d-4c3f-8b21-2d5e6f7a8b90");
    const version = doc.querySelector("OfficeApp > Version")?.textContent ?? "";
    const [major, minor] = version.split(".").map(Number);
    expect(major! * 100 + minor!).toBeGreaterThanOrEqual(104); // 1.4.0.0 or later
    expect(doc.querySelector("OfficeApp > DisplayName")?.getAttribute("DefaultValue")).toBe("Office Partner");
  });

  it("targets Word, Excel and PowerPoint", () => {
    const base = all(doc, "OfficeApp > Hosts > Host").map((host) => host.getAttribute("Name")).sort();
    expect(base).toEqual(["Document", "Presentation", "Workbook"]);
    const overrides = all(doc, "VersionOverrides > Hosts > Host").map((host) => host.getAttribute("xsi:type")).sort();
    expect(overrides).toEqual(["Document", "Presentation", "Workbook"]);
  });

  it("asks only for requirement sets every host has, and none in the overrides", () => {
    const sets = all(doc, "OfficeApp > Requirements Set").map((set) => set.getAttribute("Name"));
    expect(sets.length).toBeGreaterThan(0);
    for (const name of sets) expect(name).not.toMatch(/^(ExcelApi|WordApi|PowerPointApi)/);
    expect(all(doc, "VersionOverrides Requirements")).toHaveLength(0);
  });

  it("gives each host its own group, button and task pane ids, sharing one set of resources", () => {
    const groups = all(doc, "VersionOverrides Group").map((node) => node.getAttribute("id"));
    const buttons = all(doc, "VersionOverrides Control").map((node) => node.getAttribute("id"));
    const panes = all(doc, "VersionOverrides TaskpaneId").map((node) => node.textContent);
    for (const ids of [groups, buttons, panes]) {
      expect(ids).toHaveLength(3);
      expect(new Set(ids).size).toBe(3);
    }
    const urls = all(doc, "Resources Url").map((node) => node.getAttribute("id"));
    expect(urls).toEqual(["Taskpane.Url"]);
  });

  it("only refers to resources that exist", () => {
    const defined = new Set(all(doc, "Resources [id]").map((node) => node.getAttribute("id")));
    const used = all(doc, "VersionOverrides [resid]").map((node) => node.getAttribute("resid"));
    for (const resid of used) expect(defined.has(resid), String(resid)).toBe(true);
  });

  it("serves the pane and icons over HTTPS from an allowed app domain", () => {
    const domains = all(doc, "AppDomains > AppDomain").map((node) => node.textContent);
    const urls = [...all(doc, "[DefaultValue]").map((node) => node.getAttribute("DefaultValue")!), ...all(doc, "SourceLocation").map((node) => node.getAttribute("DefaultValue")!)]
      .filter((value) => /^https?:\/\//.test(value));
    expect(urls.length).toBeGreaterThan(0);
    for (const url of urls) {
      expect(url.startsWith("https://")).toBe(true);
      if (/\.(png|html)(\?|$)/.test(url)) expect(domains.some((domain) => url.startsWith(domain!))).toBe(true);
    }
  });
});
