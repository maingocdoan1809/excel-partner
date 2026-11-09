import { afterEach, beforeEach } from "vitest";
import { executePptTool } from "../../addin/taskpane/executor/powerpoint/index.js";
import { FakeDeck, installFakePowerPoint, type Installed } from "./fake-powerpoint.js";

export interface Fixture {
  deck: FakeDeck;
  title: string;
  agenda: string;
  data: string;
}

/** Three slides: a title slide, an agenda slide with a picture, and a slide with a table. */
export function buildDeck(api = 10): Fixture {
  const deck = new FakeDeck({ api });
  const title = deck.addSlideWith(
    [
      { name: "Title 1", placeholder: "CenterTitle", text: "Quarterly Review", top: 150, height: 100, fontSize: 40 },
      { name: "Subtitle 2", placeholder: "Subtitle", text: "Q3 results", top: 280 },
    ],
    { layoutId: "l1#" },
  );
  const agenda = deck.addSlideWith([
    { name: "Title 1", placeholder: "Title", text: "Agenda", top: 20, height: 80 },
    { name: "Content Placeholder 2", placeholder: "Body", text: "Revenue\nCosts\nOutlook", top: 120, height: 350 },
    { name: "Logo", type: "Image", left: 800, top: 20, width: 100, height: 50 },
  ]);
  const data = deck.addSlideWith([
    { name: "Title 1", placeholder: "Title", text: "Numbers", top: 20, height: 80 },
    { name: "Table 3", type: "Table", table: [["Region", "Sales"], ["EMEA", "10"], ["APAC", "12"]], left: 40, top: 120, width: 500, height: 90 },
  ]);
  return { deck, title, agenda, data };
}

/** Install a fresh fake before every test; the fixture is available through the returned getter. */
export function useFakeDeck(api = 10): () => Fixture & { installed: Installed } {
  let current: (Fixture & { installed: Installed }) | undefined;
  beforeEach(() => {
    const fixture = buildDeck(api);
    current = { ...fixture, installed: installFakePowerPoint(fixture.deck) };
  });
  afterEach(() => current?.installed.restore());
  return () => current!;
}

export async function call(tool: string, args: unknown = {}) {
  return executePptTool(tool, args);
}

/** The data of a successful call; fails the test with the error otherwise. */
export async function ok(tool: string, args: unknown = {}): Promise<any> {
  const result = await executePptTool(tool, args);
  if (!result.ok) throw new Error(`${tool} failed: ${result.error.code}: ${result.error.message}`);
  return result.data;
}

export async function fails(tool: string, args: unknown, code: string): Promise<string> {
  const result = await executePptTool(tool, args);
  if (result.ok) throw new Error(`${tool} should have failed with ${code}`);
  if (result.error.code !== code) throw new Error(`${tool} failed with ${result.error.code} (${result.error.message}), expected ${code}`);
  return result.error.message;
}
