import { resolveTheme, type ThemePreference } from "./theme-mode.js";
const THEME_KEY = "excelPartner.theme";

export function setupTheme(select: HTMLSelectElement): void {
  let preference: ThemePreference = "auto";
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === "light" || stored === "dark") preference = stored;
  } catch {}
  const system = window.matchMedia("(prefers-color-scheme: dark)");
  select.value = preference;
  const apply = () => {
    let hostTheme: Office.OfficeTheme | undefined;
    try { hostTheme = Office.context.officeTheme; } catch {}
    document.documentElement.dataset.theme = resolveTheme(preference, hostTheme, system.matches);
  };
  select.addEventListener("change", () => {
    preference = select.value === "light" || select.value === "dark" ? select.value : "auto";
    try { localStorage.setItem(THEME_KEY, preference); } catch {}
    apply();
  });
  system.addEventListener("change", apply);
  window.addEventListener("focus", apply);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) apply(); });
  const timer = window.setInterval(() => { if (!document.hidden && preference === "auto") apply(); }, 3000);
  window.addEventListener("pagehide", () => window.clearInterval(timer), { once: true });
  apply();
}
