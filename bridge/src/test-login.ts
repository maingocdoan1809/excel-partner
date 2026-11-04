import { normalizeOrigin } from "../../protocol/src/tools-testing.js";
import type { TestSessions } from "./test-sessions.js";

export interface LoginOptions {
  name: string;
  url: string;
  /** Extra sites the saved session may later be used on (for example a separate single sign-on host). */
  extraOrigins?: string[];
  store: TestSessions;
  /** Resolves when the person says they are signed in. */
  waitForPerson: () => Promise<void>;
  out: (line: string) => void;
}

/**
 * Open a real, visible browser, let the person sign in, and save the resulting session. The password is typed into the
 * site's own page and never passes through this program; only the browser's cookies and local storage are kept.
 */
export async function loginSession(options: LoginOptions): Promise<{ origins: string[] }> {
  const first = normalizeOrigin(options.url);
  if (!first) throw new Error(`'${options.url}' is not a web address (use https://host, with the port if any).`);
  const { chromium } = await import("playwright-core");
  let browser;
  const failures: string[] = [];
  for (const channel of ["chrome", "msedge"] as const) {
    try { browser = await chromium.launch({ channel, headless: false }); break; } catch (error) { failures.push(`${channel}: ${(error as Error).message.split("\n")[0]}`); }
  }
  if (!browser) throw new Error(`Could not start Chrome or Edge. ${failures.join(" | ")}`);
  try {
    const context = await browser.newContext();
    const page = await context.newPage();
    const origins = new Set<string>([first, ...(options.extraOrigins ?? []).map(item => normalizeOrigin(item)).filter((item): item is string => Boolean(item))]);
    page.on("framenavigated", frame => {
      if (frame !== page.mainFrame()) return;
      const origin = normalizeOrigin(frame.url());
      if (origin) origins.add(origin);
    });
    let closed = false;
    browser.on("disconnected", () => { closed = true; });
    await page.goto(options.url, { waitUntil: "domcontentloaded" });
    options.out(`A browser opened at ${first}. Sign in there, then come back here.`);
    await options.waitForPerson();
    if (closed) throw new Error("The browser was closed before the session was saved. Run the command again and press Enter while it is still open.");
    const state = await context.storageState();
    const list = [...origins];
    options.store.save(options.name, state, { origins: list, savedAt: new Date().toISOString(), url: options.url });
    return { origins: list };
  } finally {
    await browser.close().catch(() => undefined);
  }
}
