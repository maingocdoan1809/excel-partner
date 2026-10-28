import { z } from "zod";

/** Same shape the QA core accepts: letters/digits in 2-4 groups separated by '_' or '-'. */
export const QA_TC_ID_PATTERN = /^[A-Za-z][A-Za-z0-9]{0,11}([_-][A-Za-z0-9]{1,12}){1,3}$/;
const TcId = z.string().regex(QA_TC_ID_PATTERN).describe("Test case id such as TC_001.");
const Profile = z.string().regex(/^[a-z0-9][a-z0-9_-]{0,31}$/).describe("Profile (role) name from qa.config.json, such as admin.");
const Environment = z.string().regex(/^[a-z0-9][a-z0-9_-]{0,31}$/).optional().describe("Environment name; defaults to the project's default.");
const RunId = z.string().regex(/^\d{8}-\d{6}(-\d+)?$/).describe("Run id such as 20261028-101500; omit for the latest run.");

export const QaStatusInputSchema = z.object({ environment: Environment }).strict();
export const QaListScriptsInputSchema = z.object({}).strict();

export const QaReadScriptInputSchema = z.object({
  tcId: TcId.optional(),
  page: z.string().regex(/^[a-z][a-z0-9-]{0,40}$/).optional().describe("Page object name (automation/pages/<name>.js)."),
}).strict().refine(value => Boolean(value.tcId) !== Boolean(value.page), "Give either tcId or page.");

export const QaSaveScriptInputSchema = z.object({
  kind: z.enum(["test", "page"]).default("test"),
  tcId: TcId.optional().describe("Required for kind=test."),
  title: z.string().min(1).max(200).optional().describe("Required for kind=test: short test title (also used in the file name)."),
  name: z.string().regex(/^[a-z][a-z0-9-]{0,40}$/).optional().describe("Required for kind=page."),
  role: Profile.optional().describe("Which saved session the test uses; defaults to the first profile."),
  tags: z.array(z.string().regex(/^@?[a-z0-9_-]{1,24}$/i)).max(10).default([]),
  code: z.string().min(1).max(60_000).describe("Plain JavaScript (ES module) source. See the qa-automate skill for the required shape."),
}).strict().superRefine((value, context) => {
  if (value.kind === "test" && (!value.tcId || !value.title)) context.addIssue({ code: "custom", message: "kind=test needs tcId and title." });
  if (value.kind === "page" && !value.name) context.addIssue({ code: "custom", message: "kind=page needs name." });
});

export const QaOpenSessionInputSchema = z.object({ profile: Profile, environment: Environment }).strict();

export const QaRunInputSchema = z.object({
  tcIds: z.array(TcId).max(50).default([]).describe("Scripts to run; empty runs every approved script."),
  tag: z.string().regex(/^@?[a-z0-9_-]{1,24}$/i).optional(),
  draft: z.boolean().default(false).describe("Run unapproved scripts (supervised: the user is always asked, and the browser is visible). Requires tcIds."),
  trace: z.boolean().default(false),
  environment: Environment,
}).strict();

export const QaResultsInputSchema = z.object({ runId: RunId.optional() }).strict();

export const QaFailureInputSchema = z.object({
  tcId: TcId,
  runId: RunId.optional(),
  screenshot: z.boolean().default(true).describe("Return the failure screenshot as an attachment:// path readable with read_attached_image."),
}).strict();

export const QaSyncResultsInputSchema = z.object({
  runId: RunId.optional(),
  sheet: z.string().min(1).max(31).optional().describe("Sheet holding the test cases; defaults to the project's testCasesSheet (TestCases)."),
}).strict();

export type QaStatusInput = z.infer<typeof QaStatusInputSchema>;
export type QaReadScriptInput = z.infer<typeof QaReadScriptInputSchema>;
export type QaSaveScriptInput = z.infer<typeof QaSaveScriptInputSchema>;
export type QaOpenSessionInput = z.infer<typeof QaOpenSessionInputSchema>;
export type QaRunInput = z.infer<typeof QaRunInputSchema>;
export type QaResultsInput = z.infer<typeof QaResultsInputSchema>;
export type QaFailureInput = z.infer<typeof QaFailureInputSchema>;
export type QaSyncResultsInput = z.infer<typeof QaSyncResultsInputSchema>;

/** What the bridge sends to the task pane so it can fill the workbook; the model never handles these values. */
export const QaApplyResultsInputSchema = z.object({
  sheet: z.string().min(1).max(31),
  columns: z.record(z.string().min(1).max(60)),
  rows: z.array(z.object({
    tcId: z.string().min(1).max(40),
    values: z.record(z.union([z.string().max(2000), z.number(), z.null()])),
  }).strict()).max(5000),
}).strict();
export type QaApplyResultsInput = z.infer<typeof QaApplyResultsInputSchema>;

export const QA_TOOL_NAMES = [
  "qa_status", "qa_list_scripts", "qa_read_script", "qa_save_script", "qa_open_session", "qa_run", "qa_results", "qa_failure", "qa_sync_results",
] as const;
