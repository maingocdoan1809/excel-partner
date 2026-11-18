import { z } from "zod";

/** Environment variable names only: the add-in never stores a password. */
const EnvName = z.string().regex(/^[A-Za-z_][A-Za-z0-9_]{0,63}$/, "letters, digits and _ only");

/** What the Test screen asks for. The kit's `qa.config.json` is generated from it; the other keys of that file stay as they are. */
export const AutotestSettingsSchema = z.object({
  project: z.string().trim().min(1).max(60),
  /** `192.0.2.10` is accepted and becomes `http://192.0.2.10`. */
  url: z.preprocess(value => typeof value === "string" && value.trim() && !/^[a-z][a-z0-9+.-]*:\/\//i.test(value.trim()) ? `http://${value.trim()}` : value, z.string().trim().url().max(300)),
  testAccount: z.string().trim().max(200).default(""),
  authType: z.enum(["none", "form", "manual", "session"]).default("manual"),
  /** Name of a saved sign-in (Settings > Testing > Sign-ins). With one chosen the kit starts already signed in and never asks for a password. */
  signIn: z.string().trim().regex(/^[a-z0-9][a-z0-9_-]{0,31}$/).or(z.literal("")).default(""),
  /** false = the assistant creates and removes records named with the prefix without asking first. */
  askBeforeTestData: z.boolean().default(false),
  loginUrl: z.string().trim().max(200).default("/login"),
  usernameEnv: EnvName.default("QA_USER"),
  passwordEnv: EnvName.default("QA_PASSWORD"),
  testPrefix: z.string().trim().regex(/^[A-Za-z0-9_-]{2,30}$/).default("ZZ_AUTOTEST_"),
  talkLanguage: z.enum(["vi", "ja", "en"]).default("vi"),
  /** true = the assistant never opens the application's source code. */
  blackBox: z.boolean().default(true),
  /** The same failure this many times stops the run. */
  maxRepeat: z.number().int().min(2).max(10).default(3),
}).strict();
export type AutotestSettings = z.infer<typeof AutotestSettingsSchema>;

/** The person asks for the test kit by naming it. Nothing is guessed from other words: only this command sends the open workbook to the kit. */
export const SYSTEM_TEST_COMMAND = /^\s*\/system-test(?:\s|$)/i;

export const AUTOTEST_EVENT_KINDS = ["run_planned", "phase", "run_started", "case_started", "shot", "diagnosis", "step", "case_finished", "run_paused", "run_resumed", "run_stopped", "run_finished"] as const;
export type AutotestEventKind = (typeof AUTOTEST_EVENT_KINDS)[number];

/** One line of a run's `events.jsonl`, written by the kit. Fields beyond these are kept but not interpreted. */
export const AutotestEventSchema = z.object({
  ts: z.string(),
  kind: z.enum(AUTOTEST_EVENT_KINDS),
  suite: z.string().default(""),
  id: z.string().optional(),
  title: z.string().optional(),
  status: z.string().optional(),
  n: z.number().optional(),
  total: z.number().optional(),
  caption: z.string().optional(),
  actual: z.string().optional(),
  remark: z.string().optional(),
  reason: z.string().optional(),
  images: z.array(z.string()).optional(),
  sheet: z.string().optional(),
  row: z.number().optional(),
  header: z.number().optional(),
  cols: z.record(z.string()).optional(),
  file: z.string().optional(),
  phase: z.string().optional(),
  text: z.string().optional(),
  image: z.string().optional(),
}).passthrough();
export type AutotestEvent = z.infer<typeof AutotestEventSchema>;

/** paused = a case needs repair and the run stopped there on purpose, the rest untouched; planned = the plan is ready and nothing ran yet; preparing = the assistant is setting up or writing automation. */
export type RunPhase = "preparing" | "planned" | "running" | "paused" | "finished" | "stopped" | "idle";
export interface AutotestCaseState {
  id: string; suite: string; title: string; status: string;
  step: number; total: number; sheet: string; row: number; header: number; cols: Record<string, string>; file: string;
  remark: string; images: string[]; /** Each step keeps ITS OWN pictures (the ones taken for that step, in order). */
  steps: Array<{ n: number; status: string; caption: string; actual: string; remark: string; images: string[] }>;
  /** Pictures taken for the step in progress, which has not been recorded yet. */
  liveImages: string[];
  /** From the plan, before the case runs: the steps and expected result written in the testcase. */
  plannedSteps: string[]; expected: string; pre: string; role: string;
}
export interface AutotestRunState {
  id: string; phase: RunPhase; started: string; updated: string; stopReason: string;
  counts: Record<string, number>; cases: AutotestCaseState[]; eventCount: number;
  /** What the assistant or runner said it is doing, oldest first (last 60), and the latest one. */
  phases: Array<{ ts: string; phase: string; text: string }>; currentPhase: { phase: string; text: string } | null;
}

/** Every status a case can be in. RUNNING is the live state of a case still being tested. */
export const CASE_STATUSES = ["PLANNED", "RUNNING", "PASS", "FAIL", "BLOCKED", "PENDING", "NOT EXECUTED", "NOT AUTOMATED"] as const;
