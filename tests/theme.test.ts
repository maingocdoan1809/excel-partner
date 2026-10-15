import { describe, expect, it } from "vitest";
import { resolveTheme } from "../addin/taskpane/ui/theme-mode.js";

describe("Excel Partner appearance", () => {
  it("follows the Excel theme rather than the operating-system theme", () => {
    expect(resolveTheme("auto", { isDarkTheme: false }, true)).toBe("light");
    expect(resolveTheme("auto", { isDarkTheme: true }, false)).toBe("dark");
  });

  it("supports Office versions that provide theme colors without isDarkTheme", () => {
    expect(resolveTheme("auto", { bodyBackgroundColor: "#242424" }, false)).toBe("dark");
    expect(resolveTheme("auto", { bodyBackgroundColor: "FFFFFF" }, true)).toBe("light");
  });

  it("honors a manual preference even when Excel uses another theme", () => {
    expect(resolveTheme("dark", { isDarkTheme: false }, false)).toBe("dark");
    expect(resolveTheme("light", { isDarkTheme: true }, true)).toBe("light");
  });

  it("falls back to the system theme when Office theme information is unavailable", () => {
    expect(resolveTheme("auto", undefined, true)).toBe("dark");
    expect(resolveTheme("auto", undefined, false)).toBe("light");
    expect(resolveTheme("auto", { bodyBackgroundColor: "invalid" }, false)).toBe("light");
  });
});
