export type ThemePreference = "auto" | "light" | "dark";
export type ColorTheme = "light" | "dark";

export function resolveTheme(preference: ThemePreference, officeTheme: { isDarkTheme?: boolean; bodyBackgroundColor?: string } | undefined, prefersDark: boolean): ColorTheme {
  if (preference !== "auto") return preference;
  if (typeof officeTheme?.isDarkTheme === "boolean") return officeTheme.isDarkTheme ? "dark" : "light";
  const color = officeTheme?.bodyBackgroundColor?.replace(/^#/, "");
  if (color && /^[0-9a-f]{6}$/i.test(color)) {
    const red = parseInt(color.slice(0, 2), 16);
    const green = parseInt(color.slice(2, 4), 16);
    const blue = parseInt(color.slice(4, 6), 16);
    return (red * .2126 + green * .7152 + blue * .0722) / 255 < .5 ? "dark" : "light";
  }
  return prefersDark ? "dark" : "light";
}
