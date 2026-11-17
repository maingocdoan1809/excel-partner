import { AutotestEventSchema, type AutotestCaseState, type AutotestEvent, type AutotestRunState, type RunPhase } from "./autotest.js";

/*
 * The ONE place that turns a run's event stream into "what is happening now". The bridge (dashboard, run list) and the task pane (live view)
 * both call it, so they can never disagree about a case or a phase.
 */

/** A run that has written nothing for this long, without finishing, is shown as idle instead of running. */
export const STALE_RUN_MS = 10 * 60_000;
const MAX_STEPS_PER_CASE = 200;

/** Lines of an `events.jsonl`; a half-written or unknown line is skipped, never an error. */
export function parseEvents(text: string): AutotestEvent[] {
  const events: AutotestEvent[] = [];
  for (const line of text.split("\n")) {
    if (!line.trim()) continue;
    try {
      const parsed = AutotestEventSchema.safeParse(JSON.parse(line));
      if (parsed.success) events.push(parsed.data);
    } catch { /* a line still being written */ }
  }
  return events;
}

const blankCase = (suite: string, id: string, status = "NOT EXECUTED"): AutotestCaseState => ({
  id, suite, title: "", status, step: 0, total: 0, sheet: "", row: 0, header: 0, cols: {}, file: "", remark: "", images: [], steps: [],
  plannedSteps: [], expected: "", pre: "", role: "", liveImages: [],
});
const addOnce = (list: string[], image: string, max = 12) => (list.includes(image) ? list : [...list, image].slice(-max));
const MAX_PHASES = 60;
const str = (value: unknown) => typeof value === "string" ? value : "";

/** The one place that turns the event stream into "what is happening now". Dashboard, task pane and Excel all read this. */
export function reduceRun(id: string, events: AutotestEvent[], now = Date.now()): AutotestRunState {
  const cases = new Map<string, AutotestCaseState>();
  let phase: RunPhase = "idle";
  let started = "", updated = "", stopReason = "";
  const phases: AutotestRunState["phases"] = [];
  const get = (event: AutotestEvent) => {
    const key = `${event.suite}/${event.id}`;
    let state = cases.get(key);
    if (!state) { state = blankCase(event.suite, event.id ?? ""); cases.set(key, state); }
    if (event.sheet) { state.sheet = event.sheet; state.row = event.row ?? state.row; state.header = event.header ?? state.header; state.cols = event.cols ?? state.cols; }
    if (event.file) state.file = event.file;
    if (event.title) state.title = event.title;
    return state;
  };
  for (const event of events) {
    updated = event.ts;
    if (!started) started = event.ts;
    if (event.kind === "run_planned") {
      if (phase === "idle" || phase === "preparing") phase = "planned";
      const raw = event as Record<string, unknown>;
      for (const item of Array.isArray(raw.cases) ? raw.cases as Array<Record<string, unknown>> : []) {
        const suite = str(item.suite), id = str(item.id);
        if (!id) continue;
        const state = cases.get(`${suite}/${id}`) ?? blankCase(suite, id, "PLANNED");
        Object.assign(state, {
          title: str(item.title), expected: str(item.expected), pre: str(item.pre), role: str(item.role), file: str(item.file),
          plannedSteps: Array.isArray(item.steps) ? item.steps.map(String) : [], total: Array.isArray(item.steps) ? item.steps.length : 0,
          sheet: str(item.sheet), row: typeof item.row === "number" ? item.row : 0, header: typeof item.header === "number" ? item.header : 0,
          cols: item.cols && typeof item.cols === "object" ? item.cols as Record<string, string> : {},
        });
        cases.set(`${suite}/${id}`, state);
      }
      for (const item of Array.isArray(raw.not_automated) ? raw.not_automated as Array<Record<string, unknown>> : []) {
        const id = str(item.id);
        if (id && ![...cases.values()].some(state => state.id === id)) { const state = blankCase("", id, "NOT AUTOMATED"); state.title = str(item.title); state.remark = "No automation yet"; cases.set(`/${id}`, state); }
      }
    } else if (event.kind === "phase") {
      if (phase === "idle") phase = "preparing";
      phases.push({ ts: event.ts, phase: str(event.phase), text: str(event.text) });
      if (phases.length > MAX_PHASES) phases.shift();
    } else if (event.kind === "run_started") {
      phase = "running";
      const planned = (event as Record<string, unknown>).planned;
      if (planned && typeof planned === "object") {
        for (const [suite, ids] of Object.entries(planned as Record<string, unknown>)) {
          if (Array.isArray(ids)) for (const caseId of ids) { const key = `${suite}/${String(caseId)}`; if (!cases.has(key)) cases.set(key, blankCase(suite, String(caseId), "PLANNED")); }
        }
      }
    } else if (event.kind === "diagnosis") {
      // the picture of the screen at the moment something went wrong, next to the case it belongs to
      const image = str((event as Record<string, unknown>).image);
      if (event.id && image) { const state = get(event); state.images = addOnce(state.images, image); state.liveImages = addOnce(state.liveImages, image, 30); }
    } else if (event.kind === "shot" && event.id) {
      if (phase === "idle" || phase === "planned") phase = "running";
      const state = get(event);
      state.status = "RUNNING";
      if (event.image) { state.images = addOnce(state.images, event.image); state.liveImages = addOnce(state.liveImages, event.image, 30); }
    } else if (event.kind === "case_started" && event.id) {
      if (phase === "idle") phase = "running";
      const state = get(event);
      state.status = "RUNNING"; state.step = 0; state.steps = []; state.images = []; state.liveImages = []; state.remark = "";
    } else if (event.kind === "step" && event.id) {
      if (phase === "idle") phase = "running";
      const state = get(event);
      state.status = "RUNNING";
      state.step = event.n ?? state.step + 1;
      state.total = event.total ?? state.total;
      // the step's own pictures: the ones its event lists, else whatever was taken for it while it ran (a diagnosis picture of a crashed step)
      const stepImages = event.images?.length ? event.images : state.liveImages;
      if (state.steps.length < MAX_STEPS_PER_CASE) state.steps.push({ n: state.step, status: event.status ?? "", caption: event.caption ?? "", actual: event.actual ?? "", remark: event.remark ?? "", images: [...stepImages] });
      for (const image of stepImages) state.images = addOnce(state.images, image);
      state.liveImages = [];
    } else if (event.kind === "case_finished" && event.id) {
      const state = get(event);
      state.status = event.status ?? state.status;
      state.remark = event.remark ?? "";
    } else if (event.kind === "run_paused") {
      // stopped ON PURPOSE at the first case that needs repair; the cases not reached stay PLANNED
      phase = "paused";
      stopReason = event.reason ?? "";
    } else if (event.kind === "run_resumed") {
      // cases of THIS run are being run again (after a repair, or the ones a pause left): the run is going again, the pause is over
      phase = "running";
      stopReason = "";
    } else if (event.kind === "run_stopped") {
      phase = "stopped";
      stopReason = event.reason ?? "";
      const remaining = (event as Record<string, unknown>).remaining;
      if (Array.isArray(remaining)) {
        for (const caseId of remaining) {
          const state = get({ ...event, id: String(caseId) });
          if (state.status !== "PASS" && state.status !== "FAIL") { state.status = "NOT EXECUTED"; state.remark = `Run stopped: ${stopReason}`; }
        }
      }
    } else if (event.kind === "run_finished") {
      if (phase !== "stopped" && phase !== "paused") phase = "finished";
      if (phase === "finished") for (const state of cases.values()) if (state.status === "PLANNED") { state.status = "NOT EXECUTED"; state.remark ||= "Planned but no result"; }
    }
  }
  if ((phase === "running" || phase === "preparing") && updated && now - Date.parse(updated.replace(" ", "T")) > STALE_RUN_MS) phase = "idle";
  const list = [...cases.values()];
  const counts: Record<string, number> = {};
  for (const state of list) counts[state.status] = (counts[state.status] ?? 0) + 1;
  return { id, phase, started, updated, stopReason, counts, cases: list, eventCount: events.length, phases, currentPhase: phases.length ? { phase: phases[phases.length - 1]!.phase, text: phases[phases.length - 1]!.text } : null };
}
