import { z } from "zod";

/** Normalise "https://staging.example.com/any/path" to its origin; anything that is not plain http(s) gives null. */
export function normalizeOrigin(value: string): string | null {
  try {
    const url = new URL(value.trim());
    if (url.protocol !== "https:" && url.protocol !== "http:") return null;
    if (url.username || url.password || !url.hostname) return null;
    return url.origin;
  } catch { return null; }
}

export const TEST_SESSION_NAME_PATTERN = /^[a-z0-9][a-z0-9_-]{0,31}$/;

export const ListTestSessionsInputSchema = z.object({}).strict();

export const OpenTestBrowserInputSchema = z.object({
  session: z.string().regex(TEST_SESSION_NAME_PATTERN).optional()
    .describe("Name of a session the user saved with `excel-partner login` (see list_test_sessions). The browser starts already signed in and may reach the sites saved with it."),
  origins: z.array(z.string().min(1).max(300)).min(1).max(5).optional()
    .refine(list => !list || list.every(item => normalizeOrigin(item) !== null), "Every origin must be a plain http(s) address such as https://staging.example.com.")
    .describe("Sites the browser may open (scheme + host, port if any). Required without a session; with a session, extra sites on top of the saved ones."),
  visible: z.boolean().optional()
    .describe("Show a real browser window. Default: true without a session (the user signs in themselves), false with one (the user watches the live view)."),
}).strict().refine(value => Boolean(value.session) || Boolean(value.origins?.length), "Give a saved session or the sites the browser may open.");
export type OpenTestBrowserInput = z.infer<typeof OpenTestBrowserInputSchema>;

export const SITE_KNOWLEDGE_KINDS = ["login", "navigation", "howto", "quirk", "data"] as const;
const SiteOrigin = z.string().min(1).max(300).refine(value => normalizeOrigin(value) !== null, "Must be a plain http(s) address such as https://staging.example.com.");
const NoteName = z.string().regex(/^[a-z0-9][a-z0-9-]{0,47}$/, "lowercase letters, digits and -, up to 48 characters");

export const ReadSiteKnowledgeInputSchema = z.object({
  origin: SiteOrigin.describe("The system's address (scheme + host, port if any)."),
  name: NoteName.optional().describe("Note to read; omit to list the notes of this system (name, kind, one-line summary, date)."),
}).strict();
export type ReadSiteKnowledgeInput = z.infer<typeof ReadSiteKnowledgeInputSchema>;

export const SaveSiteKnowledgeInputSchema = z.object({
  origin: SiteOrigin,
  name: NoteName.describe("Short id, e.g. login, main-menu, create-ticket. Saving an existing name replaces that note."),
  kind: z.enum(SITE_KNOWLEDGE_KINDS).describe("login = how to get in; navigation = where things are; howto = steps for a task; quirk = surprises and workarounds; data = test data that is safe to reuse."),
  summary: z.string().min(1).max(200).describe("One line that says what the note is for; it is what future sessions see in the list."),
  content: z.string().min(1).max(6_000).describe("Markdown. Facts you verified in this system, in plain words. Never passwords, tokens, cookies or personal data."),
}).strict();
export type SaveSiteKnowledgeInput = z.infer<typeof SaveSiteKnowledgeInputSchema>;

export const DeleteSiteKnowledgeInputSchema = z.object({ origin: SiteOrigin, name: NoteName }).strict();
export type DeleteSiteKnowledgeInput = z.infer<typeof DeleteSiteKnowledgeInputSchema>;

const ResultValue = z.union([z.string().max(2000), z.number(), z.null()]);

export const RecordTestResultsInputSchema = z.object({
  sheet: z.string().min(1).max(255).describe("Name of the sheet that holds the test cases."),
  headerRow: z.number().int().min(1).max(10_000)
    .describe("1-based sheet row number of the header row directly above the data (the lowest one when headers span several rows)."),
  results: z.array(z.object({
    row: z.number().int().min(2).max(1_048_576).describe("1-based sheet row number to write into (the test case row, or a step row)."),
    values: z.record(ResultValue).describe("Header text -> value. Headers are matched to the header row ignoring case and extra spaces."),
  }).strict()).min(1).max(200),
  addMissingColumns: z.boolean().default(true)
    .describe("Append a column after the last header for each header that does not exist yet (copies the look of the neighbouring header)."),
  allowOverwrite: z.boolean().default(false)
    .describe("Only after the user agreed: overwrite non-empty cells that this tool did not write itself in this session."),
}).strict();
export type RecordTestResultsInput = z.infer<typeof RecordTestResultsInputSchema>;

/** Test cases whose script is worth keeping: the page behaved as expected, or the application is wrong. */
export const SCRIPT_OUTCOMES = ["pass", "fail_app_defect"] as const;
export type ScriptOutcome = (typeof SCRIPT_OUTCOMES)[number];
/** Every way a tested case can end; only the first two leave a regression script behind. */
export const CASE_OUTCOMES = ["pass", "fail_app_defect", "fail_other", "blocked", "needs_review"] as const;
export type CaseOutcome = (typeof CASE_OUTCOMES)[number];
export const CASE_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;

export const SaveTestScriptInputSchema = z.object({
  origin: SiteOrigin.describe("The system under test (scheme + host, port if any)."),
  caseId: z.string().regex(CASE_ID_PATTERN, "letters, digits, . _ - up to 64 characters").describe("The test case's ID from the sheet (for example TC-012). Saving the same ID again replaces the script."),
  title: z.string().min(1).max(200).describe("What the test case checks, in a few words."),
  sheet: z.string().min(1).max(255).describe("Sheet that holds this test case."),
  row: z.number().int().min(2).max(1_048_576).describe("1-based sheet row of the test case (the row you record its result in)."),
  outcome: z.enum(SCRIPT_OUTCOMES).describe("pass = the application behaved as the test case expects. fail_app_defect = you are sure the APPLICATION is wrong. Never use this when you were blocked by a technical problem or did not understand the system: such cases get no script."),
  evidence: z.string().min(20).max(2_000).describe("What you saw on the page. For fail_app_defect: the expected result, the actual result, and why this is the application's fault and not yours."),
  script: z.string().min(40).max(20_000).describe("A complete Playwright Test file (JavaScript) that repeats the case from the first step with the same data and asserts the expected result with expect(). See the tool description for the required shape."),
}).strict();
export type SaveTestScriptInput = z.infer<typeof SaveTestScriptInputSchema>;

export const FinishTestRunInputSchema = z.object({
  sheet: z.string().min(1).max(255),
  cases: z.array(z.object({
    row: z.number().int().min(2).max(1_048_576).describe("Sheet row of the test case, as used in record_test_results."),
    caseId: z.string().regex(CASE_ID_PATTERN).describe("The test case's ID from the sheet."),
    outcome: z.enum(CASE_OUTCOMES).describe("How the case ended. pass and fail_app_defect need a script saved with save_test_script; the others must not have one."),
  }).strict()).min(1).max(200),
}).strict();
export type FinishTestRunInput = z.infer<typeof FinishTestRunInputSchema>;
