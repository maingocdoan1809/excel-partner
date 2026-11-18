import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import type { AutotestEvent } from "../protocol/src/autotest.js";
import { AutotestHome, AutotestHomeError, parseKnowledge } from "../bridge/src/autotest-home.js";
import { AutotestRuns, parseEvents, reduceRun, STALE_RUN_MS } from "../bridge/src/autotest-runs.js";
import { AutotestRoutes } from "../bridge/src/routes/autotest.js";
import { SYSTEM_TEST_COMMAND } from "../protocol/src/autotest.js";
import { ProgressWriter } from "../frontend/taskpane/executor/autotest-sheet.js";

const temps: string[] = [];
const temp = () => { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "autotest-test-")); temps.push(dir); return dir; };
afterEach(() => { for (const dir of temps.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });

/** Wait until something is true (a slow machine only makes this longer, never wrong). */
async function until(condition: () => boolean, ms = 5000) {
  const end = Date.now() + ms;
  while (!condition() && Date.now() < end) await new Promise(resolve => setTimeout(resolve, 25));
  expect(condition()).toBe(true);
}

const line = (event: Record<string, unknown>) => JSON.stringify({ ts: "2026-11-17 10:00:00", suite: "orders", ...event }) + "\n";
const E = (event: Record<string, unknown>) => ({ ts: "2026-11-17 10:00:00", suite: "orders", ...event }) as AutotestEvent;
const where = { sheet: "Cases", row: 5, header: 3, cols: { result: "Result", date: "Date" }, file: "/home/x/testcases/Orders.xlsx" };

describe("run state from the event stream", () => {
  it("skips half-written and unknown lines", () => {
    const text = line({ kind: "case_started", id: "A" }) + "{\"kind\":\"case_fin" + "\n" + line({ kind: "nonsense" });
    expect(parseEvents(text)).toHaveLength(1);
  });

  it("follows a case from running to finished and counts statuses", () => {
    const state = reduceRun("r1", [
      E({ kind: "run_started", planned: { orders: ["A", "B", "C"] } }),
      E({ kind: "case_started", id: "A", ...where }),
      E({ kind: "step", id: "A", n: 1, total: 2, status: "PASS", caption: "open", actual: "ok", images: ["/o/1.png"] }),
      E({ kind: "case_finished", id: "A", status: "PASS" }),
      E({ kind: "case_started", id: "B" }),
    ], Date.parse("2026-11-17T10:00:30"));
    expect(state.phase).toBe("running");
    expect(state.counts).toEqual({ PASS: 1, RUNNING: 1, PLANNED: 1 });
    const a = state.cases.find(item => item.id === "A")!;
    expect(a).toMatchObject({ status: "PASS", sheet: "Cases", row: 5, header: 3, images: ["/o/1.png"] });
  });

  it("a stopped run keeps its reason and marks the cases that never ran", () => {
    const state = reduceRun("r1", [
      E({ kind: "run_started", planned: { orders: ["A", "B", "C"] } }),
      E({ kind: "case_started", id: "A" }), E({ kind: "case_finished", id: "A", status: "FAIL" }),
      E({ kind: "run_stopped", reason: "the same error repeated 3 times: login 500", remaining: ["B", "C"] }),
    ]);
    expect(state.phase).toBe("stopped");
    expect(state.stopReason).toContain("repeated 3 times");
    expect(state.cases.filter(item => item.status === "NOT EXECUTED").map(item => item.id)).toEqual(["B", "C"]);
  });

  it("a run that went quiet without finishing is idle, not running", () => {
    const events = [E({ kind: "case_started", id: "A" })];
    expect(reduceRun("r", events, Date.parse("2026-11-17T10:00:00") + STALE_RUN_MS + 1000).phase).toBe("idle");
    expect(reduceRun("r", events, Date.parse("2026-11-17T10:00:05")).phase).toBe("running");
  });
});

describe("runs on disk", () => {
  it("lists runs, reads state, and delivers only new lines to a watcher", async () => {
    const out = temp();
    const dir = path.join(out, "runs", "20261117_run");
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, "events.jsonl"), line({ kind: "run_started", planned: { orders: ["A"] } }));
    const runs = new AutotestRuns(out);
    expect(runs.list().map(item => item.id)).toEqual(["20261117_run"]);
    const seen: string[] = [];
    const stop = runs.watch((run, event) => seen.push(`${run}:${event.kind}`), 20);
    fs.appendFileSync(path.join(dir, "events.jsonl"), line({ kind: "case_started", id: "A" }) + "{\"kind\":\"step\"");   // second line still being written
    await until(() => seen.includes("20261117_run:case_started"));
    expect(seen).toEqual(["20261117_run:case_started"]);              // the half-written line is not delivered yet
    fs.appendFileSync(path.join(dir, "events.jsonl"), ",\"id\":\"A\",\"suite\":\"orders\",\"ts\":\"2026-11-17 10:00:01\"}\n");
    await until(() => seen.length === 2);
    stop();
    expect(seen).toEqual(["20261117_run:case_started", "20261117_run:step"]);
    expect(runs.state("20261117_run")!.cases[0]!.status).toBe("RUNNING");
  });

  it("serves an image only from inside the output folder", () => {
    const out = temp();
    const dir = path.join(out, "runs", "r1");
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, "events.jsonl"), line({ kind: "case_started", id: "A" }));
    fs.writeFileSync(path.join(dir, "a.png"), "x");
    const secret = path.join(temp(), "secret.png");
    fs.writeFileSync(secret, "x");
    const runs = new AutotestRuns(out);
    expect(runs.image("r1", path.join(dir, "a.png"))).toBe(fs.realpathSync(path.join(dir, "a.png")));
    expect(runs.image("r1", secret)).toBeNull();
    // an event written by a run with a relative QA_OUT names the picture relative to the autotest folder (the parent of the output folder)
    expect(runs.image("r1", path.join(path.basename(out), "runs", "r1", "a.png"))).toBe(fs.realpathSync(path.join(dir, "a.png")));
    expect(runs.image("r1", path.join(path.basename(out), "..", "..", "x.png"))).toBeNull();
    expect(runs.image("r1", path.join(dir, "..", "..", "..", "x.png"))).toBeNull();
    expect(runs.image("nope", path.join(dir, "a.png"))).toBeNull();
  });
});

describe("the one autotest folder", () => {
  function fixture() {
    const root = temp();
    const kit = path.join(root, "kit-src");
    fs.mkdirSync(path.join(kit, "kit", "qakit"), { recursive: true });
    fs.mkdirSync(path.join(kit, "kit", "project"), { recursive: true });
    fs.mkdirSync(path.join(kit, "skills", "system-test", "references"), { recursive: true });
    fs.writeFileSync(path.join(kit, "VERSION"), "0.2.0");
    fs.writeFileSync(path.join(kit, "kit", "qakit", "__init__.py"), "# v1");
    fs.writeFileSync(path.join(kit, "kit", "project", "hooks.py"), "# template");
    fs.writeFileSync(path.join(kit, "kit", "requirements.txt"), "playwright\n");
    fs.writeFileSync(path.join(kit, "kit", "qa.config.example.json"), JSON.stringify({ project: "x", app: { url: "http://a" }, report: { result_values: { PASS: "OK" } }, custom: { keep: true } }));
    fs.writeFileSync(path.join(kit, "skills", "system-test", "SKILL.md"), "---\nname: system-test\n---\nbody");
    fs.writeFileSync(path.join(kit, "skills", "system-test", "references", "addin.md"), "addin");
    const skills = path.join(root, "skills");
    const home = new AutotestHome(path.join(root, "pointer.json"), path.join(root, "home"), kit, skills);
    return { root, kit, skills, home };
  }
  const settings = { project: "Shop", url: "https://shop.test", testAccount: "qa@shop.test", authType: "form", usernameEnv: "SHOP_USER", passwordEnv: "SHOP_PASS", maxRepeat: 4 };

  it("writes qa.config.json from the settings, keeps the other keys and installs kit and skills", () => {
    const { home, skills } = fixture();
    const saved = home.save(settings);
    expect(saved.installed).toEqual(expect.arrayContaining(["qakit/__init__.py", "project/hooks.py", "knowledge/index.md", "system-test/SKILL.md", "system-test/references/addin.md"]));
    const config = JSON.parse(fs.readFileSync(home.configFile, "utf8"));
    expect(config).toMatchObject({ project: "Shop", app: { url: "https://shop.test", test_account: "qa@shop.test" }, auth: { type: "form", username_env: "SHOP_USER", password_env: "SHOP_PASS" }, run: { max_repeat: 4 }, custom: { keep: true } });
    expect(JSON.stringify(config)).not.toMatch(/password"\s*:\s*"[^"]+"/i);
    expect(home.settings()).toMatchObject({ project: "Shop", maxRepeat: 4, blackBox: true });
    expect(home.resultValues()).toEqual({ PASS: "OK" });
    expect(fs.existsSync(path.join(skills, "system-test", "SKILL.md"))).toBe(true);
  });

  it("refreshes framework code but never overwrites the project's own hooks", () => {
    const { home, kit } = fixture();
    home.save(settings);
    fs.writeFileSync(path.join(home.dir, "project", "hooks.py"), "# mine");
    fs.writeFileSync(path.join(kit, "kit", "qakit", "__init__.py"), "# v2");
    fs.writeFileSync(path.join(kit, "kit", "project", "hooks.py"), "# template v2");
    const again = home.installKit();
    expect(again).toContain("qakit/__init__.py");
    expect(again).not.toContain("project/hooks.py");
    expect(fs.readFileSync(path.join(home.dir, "project", "hooks.py"), "utf8")).toBe("# mine");
    expect(fs.readFileSync(path.join(home.dir, "qakit", "__init__.py"), "utf8")).toBe("# v2");
  });

  it("rejects a bad setting with every reason, and a password-like value is not a field at all", () => {
    const { home } = fixture();
    expect(() => home.save({ ...settings, url: "not a url", usernameEnv: "bad name" })).toThrow(AutotestHomeError);
    expect(() => home.save({ ...settings, password: "secret" })).toThrow(AutotestHomeError);
  });

  it("lets the user choose another folder once, but not a root or somebody else's folder", () => {
    const { home, root } = fixture();
    const other = path.join(root, "chosen");
    home.choose(other);
    expect(home.dir).toBe(other);
    expect(() => home.choose(path.parse(root).root)).toThrow(AutotestHomeError);
    const busy = path.join(root, "busy");
    fs.mkdirSync(busy);
    fs.writeFileSync(path.join(busy, "notes.txt"), "mine");
    expect(() => home.choose(busy)).toThrow(/not empty/);
  });

  it("keeps the open workbook as the testcase source and remembers which one is current", () => {
    const { home } = fixture();
    home.save(settings);
    const xlsx = Buffer.concat([Buffer.from("PK"), Buffer.alloc(200, 1)]);
    const file = home.saveWorkbook("../Orders v2.xlsx", xlsx);
    expect(path.basename(file)).toBe("Orders v2.xlsx");
    expect(path.dirname(file)).toBe(home.testcasesDir);
    expect(home.activeWorkbook()).toBe("Orders v2.xlsx");
    expect(() => home.saveWorkbook("x.xlsx", Buffer.from("not a zip at all, long enough to pass the size check ".repeat(5)))).toThrow(/not an Excel/);
  });

  it("reads the knowledge index one line per fact", () => {
    const lines = parseKnowledge("# head\n- [screen:orders] [kind:error] [tag:timeout] Save stays off until a customer is chosen -> pages/orders.py\nnoise\n- [kind:quirk] global fact\n");
    expect(lines).toEqual([
      { screen: "orders", kind: "error", tags: ["timeout"], fact: "Save stays off until a customer is chosen", ref: "pages/orders.py", line: 2 },
      { screen: "", kind: "quirk", tags: [], fact: "global fact", ref: "-", line: 4 },
    ]);
  });
});

describe("the dashboard", () => {
  async function serve() {
    const root = temp();
    const home = new AutotestHome(path.join(root, "p.json"), path.join(root, "home"), path.join(root, "none"));
    const dir = path.join(home.outDir, "runs", "r1");
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, "events.jsonl"), line({ kind: "case_started", id: "A", title: "<img src=x onerror=alert(1)>" }));
    const routes = new AutotestRoutes({ sessions: {} as never, home, runtime: {} as never });
    const server = http.createServer((req, res) => { void routes.dashboard(req, res, new URL(req.url ?? "/", "http://localhost")); });
    await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
    const base = `http://127.0.0.1:${(server.address() as { port: number }).port}`;
    return { base, close: () => new Promise<void>(resolve => server.close(() => resolve())) };
  }

  it("hands the page a token, and the API refuses calls without it", async () => {
    const { base, close } = await serve();
    try {
      const page = await (await fetch(`${base}/tests`)).text();
      const token = /var T="([0-9a-f]+)"/.exec(page)?.[1];
      expect(token).toBeTruthy();
      expect((await fetch(`${base}/tests/api/runs`)).status).toBe(401);
      expect((await fetch(`${base}/tests/api/runs?t=wrong`)).status).toBe(401);
      const listed = await (await fetch(`${base}/tests/api/runs?t=${token}`)).json() as { runs: Array<{ id: string }> };
      expect(listed.runs.map(item => item.id)).toEqual(["r1"]);
      const run = await (await fetch(`${base}/tests/api/run?t=${token}&id=r1`)).json() as { run: { cases: Array<{ title: string }> } };
      expect(run.run.cases[0]!.title).toContain("<img");   // data is delivered as data; the page shows it with textContent
      expect(page).not.toContain("innerHTML");
      expect((await fetch(`${base}/tests/api/image?t=${token}&id=r1&path=${encodeURIComponent("/etc/passwd")}`)).status).toBe(404);
    } finally { await close(); }
  });
});

describe("progress into the testcase sheet", () => {
  it("writes into the sheet row the case lives in, merges quick updates, and ignores other files", async () => {
    const writes: unknown[] = [];
    const writer = new ProgressWriter(() => ({ PASS: "OK" }), message => { throw new Error(message); }, 5, (async (input: unknown) => { writes.push(input); }) as never);
    writer.file = where.file;
    expect(writer.push(E({ kind: "case_started", id: "A", ...where }))).toBe(true);
    expect(writer.push(E({ kind: "step", id: "A", n: 1, total: 2, ...where }))).toBe(true);
    expect(writer.push(E({ kind: "case_started", id: "Z", ...where, file: "/other/File.xlsx" }))).toBe(false);
    expect(writer.push(E({ kind: "case_finished", id: "A", status: "PASS", remark: "", ...where }))).toBe(true);
    await writer.flush();
    expect(writes).toHaveLength(1);
    const call = writes[0] as { sheet: string; headerRow: number; allowOverwrite: boolean; results: Array<{ row: number; values: Record<string, string> }> };
    expect(call).toMatchObject({ sheet: "Cases", headerRow: 3, allowOverwrite: true });
    expect(call.results).toHaveLength(1);
    expect(call.results[0]!.row).toBe(5);
    expect(call.results[0]!.values.Result).toBe("OK");
    expect(call.results[0]!.values.Date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
  });
});

describe("the kit's circuit breaker (Python)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("stops on the third identical failure, ignores numbers and quoted values, and never counts a human-needed case", () => {
    const code = [
      "import json",
      "from qakit.core.breaker import Breaker",
      "b = Breaker()",
      "fail = lambda n, q: [dict(status='FAIL', actual=f'Toast 「{q}」 not shown after {n} ms')]",
      "out = [b.observe('FAIL', fail(100, 'a')), b.observe('FAIL', fail(250, 'b')), Breaker().observe('PENDING', fail(1, 'c')), b.observe('FAIL', fail(999, 'd'))]",
      "other = Breaker(); out.append(other.observe('FAIL', [dict(status='FAIL', actual='login broke')])); out.append(other.observe('FAIL', [dict(status='FAIL', actual='menu missing')]))",
      "print(json.dumps(out))",
    ].join("\n");
    const run = spawnSync(python, ["-c", code], { cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: temp(), QA_MAX_REPEAT: "3" }, encoding: "utf8" });
    expect(run.status, run.stderr).toBe(0);
    const out = JSON.parse(run.stdout) as Array<string | null>;
    expect([out[0], out[1], out[2]]).toEqual([null, null, null]);
    expect(out[3]).toMatch(/repeated 3 times/);
    expect([out[4], out[5]]).toEqual([null, null]);   // two different errors never add up
  });
});

describe("the plan and the phases are visible before anything runs", () => {
  const planned = E({ kind: "run_planned", run_id: "r1", suite: "", cases: [
    { suite: "orders", id: "A", title: "Create order", pre: "logged in", expected: "toast Saved", steps: ["open", "click New order"], role: "", sheet: "Cases", row: 3, header: 2, cols: { result: "Result" }, file: "/x/Orders.xlsx" },
    { suite: "orders", id: "B", title: "Export", pre: "", expected: "csv", steps: ["click Export"], sheet: "Cases", row: 4, header: 2, cols: {}, file: "/x/Orders.xlsx" },
  ], not_automated: [{ id: "C", title: "Import" }] });

  it("shows every planned case with its steps and expected result, and the phase the assistant reported", () => {
    const state = reduceRun("r1", [E({ kind: "phase", phase: "automate", text: "writing 2 cases", suite: "" }), planned]);
    expect(state.phase).toBe("planned");
    expect(state.currentPhase).toEqual({ phase: "automate", text: "writing 2 cases" });
    expect(state.counts).toEqual({ PLANNED: 2, "NOT AUTOMATED": 1 });
    expect(state.cases.find(item => item.id === "A")).toMatchObject({ title: "Create order", expected: "toast Saved", plannedSteps: ["open", "click New order"], total: 2, sheet: "Cases", row: 3 });
  });

  it("evidence pictures appear the moment they are taken, before the step ends, and nothing planned is lost when the run starts", () => {
    const state = reduceRun("r1", [planned, E({ kind: "run_started", planned: { orders: ["A", "B"] } }), E({ kind: "case_started", id: "A" }),
      E({ kind: "shot", id: "A", n: 1, caption: "click", image: "/o/a1.png" })], Date.parse("2026-11-17T10:00:05"));
    const a = state.cases.find(item => item.id === "A")!;
    expect(state.phase).toBe("running");
    expect(a).toMatchObject({ status: "RUNNING", images: ["/o/a1.png"], plannedSteps: ["open", "click New order"] });
    expect(state.cases.find(item => item.id === "B")!.status).toBe("PLANNED");
  });

  it("a finished run turns cases that never got a result into NOT EXECUTED", () => {
    const state = reduceRun("r1", [planned, E({ kind: "run_started" }), E({ kind: "case_finished", id: "A", status: "PASS" }), E({ kind: "run_finished" })]);
    expect(state.cases.map(item => item.status)).toEqual(["PASS", "NOT EXECUTED", "NOT AUTOMATED"]);
  });

});

describe("kit: plan events and status line (Python)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("`status` writes a phase line to the open run, and to the ad-hoc folder once that run has ended", () => {
    const home = temp();
    fs.writeFileSync(path.join(home, "qa.config.json"), JSON.stringify({ project: "t" }));
    const runDir = path.join(home, "test_execution_output", "runs", "r1");
    fs.mkdirSync(runDir, { recursive: true });
    fs.writeFileSync(path.join(runDir, "events.jsonl"), line({ kind: "run_planned", cases: [] }));
    const say = (phase: string, text: string) => spawnSync(python, ["-m", "qakit.tools.status", phase, text], { cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: home, QA_OUT: "", QA_EVENTS: "" }, encoding: "utf8" });
    expect(say("plan", "waiting for the user").status).toBe(0);
    expect(fs.readFileSync(path.join(runDir, "events.jsonl"), "utf8")).toContain("waiting for the user");
    fs.appendFileSync(path.join(runDir, "events.jsonl"), line({ kind: "run_finished" }));
    expect(say("setup", "next job").status).toBe(0);
    expect(fs.readFileSync(path.join(runDir, "events.jsonl"), "utf8")).not.toContain("next job");
    const adhoc = path.join(home, "test_execution_output", "adhoc");
    expect(fs.readdirSync(adhoc).length).toBe(1);
    expect(fs.readFileSync(path.join(adhoc, fs.readdirSync(adhoc)[0]!, "events.jsonl"), "utf8")).toContain("next job");
    expect(say("nonsense", "x").status).not.toBe(0);
  });
});

describe("signing in without a password", () => {
  function homeWith(states: Record<string, string>) {
    const root = temp();
    fs.mkdirSync(path.join(root, "kit-src", "kit", "qakit"), { recursive: true });
    fs.writeFileSync(path.join(root, "kit-src", "kit", "qakit", "__init__.py"), "");
    fs.writeFileSync(path.join(root, "kit-src", "kit", "qa.config.example.json"), JSON.stringify({ project: "x", browser: { mode: "cdp" } }));
    fs.writeFileSync(path.join(root, "kit-src", "VERSION"), "1");
    const source = { list: () => Object.keys(states).map(name => ({ name, url: "http://a", savedAt: "" })), statePath: (name: string) => states[name] ?? null };
    return new AutotestHome(path.join(root, "p.json"), path.join(root, "home"), path.join(root, "kit-src"), null, source);
  }
  const base = { project: "Shop", url: "https://shop.test", authType: "session" };

  it("points the kit at the saved sign-in, forces its own browser, and never writes a secret", () => {
    const home = homeWith({ admin: "/state/admin.state.json" });
    home.save({ ...base, signIn: "admin", askBeforeTestData: false });
    const config = JSON.parse(fs.readFileSync(home.configFile, "utf8"));
    expect(config.auth).toMatchObject({ type: "session", session_name: "admin", storage_state: "/state/admin.state.json", logged_in: { url_not_contains: "/login" } });
    expect(config.browser.mode).toBe("launch");
    expect(config.run).toMatchObject({ ask_before_test_data: false });
    expect(home.settings()).toMatchObject({ authType: "session", signIn: "admin", askBeforeTestData: false });
    expect(home.signInList().map(item => item.name)).toEqual(["admin"]);
  });

  it("refuses a sign-in that does not exist, or none at all, with a way forward", () => {
    const home = homeWith({ admin: "/state/admin.state.json" });
    expect(() => home.save({ ...base, signIn: "ghost" })).toThrow(/not found/);
    expect(() => home.save({ ...base, signIn: "" })).toThrow(/Choose a saved sign-in/);
  });

  it("remembers that the person wants to be asked before test records are created", () => {
    const home = homeWith({ admin: "/state/admin.state.json" });
    home.save({ ...base, signIn: "admin", askBeforeTestData: true });
    expect(home.settings()!.askBeforeTestData).toBe(true);
  });
});

describe("what the screen said when something went wrong", () => {
  it("puts the diagnosis picture next to the case it belongs to", () => {
    const state = reduceRun("r", [E({ kind: "case_started", id: "A" }), E({ kind: "diagnosis", id: "A", image: "/o/diagnose/a.png", summary: "Customer Name required" })]);
    expect(state.cases[0]!.images).toEqual(["/o/diagnose/a.png"]);
  });
});

describe("saving the environment form", () => {
  function plain() {
    const root = temp();
    fs.mkdirSync(path.join(root, "k", "kit", "qakit"), { recursive: true });
    fs.writeFileSync(path.join(root, "k", "kit", "qakit", "__init__.py"), "");
    fs.writeFileSync(path.join(root, "k", "kit", "qa.config.example.json"), "{}");
    return new AutotestHome(path.join(root, "p.json"), path.join(root, "home"), path.join(root, "k"), null, { list: () => [], statePath: () => null });
  }
  const good = { project: "Shop", url: "http://192.0.2.10", authType: "manual" };

  it("completes an address typed without http://", () => {
    const home = plain();
    home.save({ ...good, url: "192.0.2.10" });
    expect(home.settings()!.url).toBe("http://192.0.2.10");
  });

  it("does not need the variable-name boxes unless the kit fills the form", () => {
    expect(() => plain().save(good)).not.toThrow();      // not sent at all: defaults apply
  });

  it("tells a person who typed the user name or password into a variable-name box exactly what that box wants", () => {
    const home = plain();
    const say = (patch: object) => { try { home.save({ ...good, authType: "form", ...patch }); return ""; } catch (error) { return (error as Error).message; } };
    const message = say({ usernameEnv: "admin@shop.vn", passwordEnv: "Abc@123" });
    expect(message).toContain("NAME of a variable");
    expect(message).toContain("Do not type a password here");
    expect(message).not.toMatch(/Invalid|letters, digits and _ only/);
    expect(say({ testPrefix: "ZZ AUTO" })).toContain("no spaces");
    expect(say({ maxRepeat: 0 })).toContain("from 2 to 10");
  });
});

describe("kit: writing events (Python)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("writes an event that has its own `file`, `kind` and `suite` fields, and never raises", () => {
    const out = temp();
    const code = [
      "from qakit.core import events",
      "events.emit('case_started', id='A', file='/x/Orders.xlsx', sheet='Cases', row=3)",
      "events.emit('step', id='A', kind='odd', suite='orders', n=1)",
      "events.emit_to(r'" + path.join(out, "other.jsonl").replace(/\\/g, "/") + "', 'phase', phase='fix', text='x', file='f')",
      "print('done')",
    ].join("\n");
    const run = spawnSync(python, ["-c", code], { cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: temp(), QA_EVENTS: path.join(out, "events.jsonl") }, encoding: "utf8" });
    expect(run.status, run.stderr).toBe(0);
    const events = fs.readFileSync(path.join(out, "events.jsonl"), "utf8").trim().split("\n").map(item => JSON.parse(item));
    expect(events[0]).toMatchObject({ kind: "case_started", id: "A", file: "/x/Orders.xlsx", row: 3 });
    expect(events).toHaveLength(2);
    expect(fs.readFileSync(path.join(out, "other.jsonl"), "utf8")).toContain("\"file\": \"f\"");
  });
});

describe("a run that pauses at the first case needing repair", () => {
  const planned = E({ kind: "run_planned", cases: [{ suite: "orders", id: "A", steps: ["x"] }, { suite: "orders", id: "B", steps: ["x"] }, { suite: "orders", id: "C", steps: ["x"] }] });

  it("shows paused with the reason, keeps the cases not reached as PLANNED, and stays paused when the runner then closes the run", () => {
    const events = [planned, E({ kind: "run_started" }), E({ kind: "case_started", id: "A" }), E({ kind: "case_finished", id: "A", status: "PASS" }),
      E({ kind: "case_started", id: "B" }), E({ kind: "case_finished", id: "B", status: "NOT EXECUTED" }),
      E({ kind: "run_paused", reason: "case B needs repair: NOT EXECUTED: save refused", remaining: ["C"] }), E({ kind: "run_finished" })];
    const state = reduceRun("r", events);
    expect(state.phase).toBe("paused");
    expect(state.stopReason).toContain("case B needs repair");
    expect(state.cases.map(item => item.status)).toEqual(["PASS", "NOT EXECUTED", "PLANNED"]);
  });


  it("the add-in turns on repair-at-once for the kit", () => {
    const root = temp();
    fs.mkdirSync(path.join(root, "k", "kit", "qakit"), { recursive: true });
    fs.writeFileSync(path.join(root, "k", "kit", "qakit", "__init__.py"), "");
    fs.writeFileSync(path.join(root, "k", "kit", "qa.config.example.json"), "{}");
    const home = new AutotestHome(path.join(root, "p.json"), path.join(root, "home"), path.join(root, "k"));
    home.save({ project: "P", url: "http://a", authType: "manual" });
    expect(JSON.parse(fs.readFileSync(home.configFile, "utf8")).run).toMatchObject({ stop_on_problem: true, case_retries: 1, max_attempts: 2 });
  });
});

describe("the kit on this computer follows the release without anybody pressing Save", () => {
  function setup() {
    const root = temp();
    const kit = path.join(root, "kit-src");
    fs.mkdirSync(path.join(kit, "kit", "qakit"), { recursive: true });
    fs.mkdirSync(path.join(kit, "kit", "project"), { recursive: true });
    fs.mkdirSync(path.join(kit, "skills", "system-test"), { recursive: true });
    fs.writeFileSync(path.join(kit, "kit", "qakit", "__init__.py"), "# v1");
    fs.writeFileSync(path.join(kit, "kit", "qakit", "watch.py"), "# watch v1");
    fs.writeFileSync(path.join(kit, "kit", "project", "hooks.py"), "# template");
    fs.writeFileSync(path.join(kit, "kit", "qa.config.example.json"), "{}");
    fs.writeFileSync(path.join(kit, "skills", "system-test", "SKILL.md"), "skill v1");
    const skills = path.join(root, "skills");
    const make = () => new AutotestHome(path.join(root, "p.json"), path.join(root, "home"), kit, skills);
    // `fresh()` = the bridge restarted with a new release: the release does not change while the bridge runs, so the fingerprint is read once
    return { root, kit, skills, home: make(), fresh: make };
  }
  const saved = { project: "P", url: "http://a", authType: "manual" };

  it("does nothing before the person has set the kit up", () => {
    expect(setup().home.syncKit()).toMatchObject({ updated: [], skipped: "not set up yet" });
  });

  it("is a no-op while the kit is current, and refreshes old code and skills (keeping a copy) when the release has changed", () => {
    const { home, kit, skills, fresh } = setup();
    home.save(saved);
    expect(home.syncKit().skipped).toBe("already current");
    // a release with new kit code and a new skill text; the person's hooks are theirs
    fs.writeFileSync(path.join(kit, "kit", "qakit", "watch.py"), "# watch v2");
    fs.writeFileSync(path.join(kit, "kit", "qakit", "diagnose.py"), "# new file");
    fs.writeFileSync(path.join(kit, "skills", "system-test", "SKILL.md"), "skill v2");
    fs.writeFileSync(path.join(home.dir, "project", "hooks.py"), "# mine");
    const synced = fresh().syncKit();
    expect(synced.updated).toEqual(expect.arrayContaining(["qakit/watch.py", "qakit/diagnose.py", "system-test/SKILL.md"]));
    expect(fs.readFileSync(path.join(home.dir, "qakit", "watch.py"), "utf8")).toBe("# watch v2");
    expect(fs.readFileSync(path.join(skills, "system-test", "SKILL.md"), "utf8")).toBe("skill v2");
    expect(fs.readFileSync(path.join(home.dir, "project", "hooks.py"), "utf8")).toBe("# mine");
    expect(synced.backup).toMatch(/^qakit_backup\/\d{14}$/);
    expect(fs.readFileSync(path.join(home.dir, synced.backup!, "qakit", "watch.py"), "utf8")).toBe("# watch v1");
    expect(fs.readFileSync(path.join(home.dir, synced.backup!, "skills", "system-test", "SKILL.md"), "utf8")).toBe("skill v1");
    expect(fresh().syncKit().skipped).toBe("already current");
  });

  it("waits while a run is going", () => {
    const { home, kit, fresh } = setup();
    home.save(saved);
    const dir = path.join(home.outDir, "runs", "r1");
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, "events.jsonl"), JSON.stringify({ ts: new Date(Date.now() - new Date().getTimezoneOffset() * 60_000).toISOString().slice(0, 19).replace("T", " "), kind: "run_started", suite: "" }) + "\n");
    fs.writeFileSync(path.join(kit, "kit", "qakit", "watch.py"), "# watch v2");
    expect(fresh().syncKit()).toMatchObject({ updated: [], skipped: "a run is in progress" });
    expect(fs.readFileSync(path.join(home.dir, "qakit", "watch.py"), "utf8")).toBe("# watch v1");
  });

  it("adds the repair settings an older config lacks, but never overrides a choice", () => {
    const { home, kit, fresh } = setup();
    home.save(saved);
    const config = JSON.parse(fs.readFileSync(home.configFile, "utf8"));
    delete config.run.stop_on_problem; config.run.case_retries = 0;      // an old config, and one deliberate choice
    fs.writeFileSync(home.configFile, JSON.stringify(config));
    fs.writeFileSync(path.join(kit, "kit", "qakit", "watch.py"), "# watch v2");
    fresh().syncKit();
    expect(JSON.parse(fs.readFileSync(home.configFile, "utf8")).run).toMatchObject({ stop_on_problem: true, case_retries: 0, max_attempts: 2 });
  });
});

describe("kit: the red box (Python + Pillow)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const hasPillow = (() => { try { return spawnSync(python, ["-c", "import PIL"]).status === 0; } catch { return false; } })();

  it.skipIf(!hasPillow)("a box around an element as wide as the screen still shows on all four sides", () => {
    const out = path.join(temp(), "frame.png").replace(/\\/g, "/");
    const code = [
      "from PIL import Image", "from qakit.core import evidence",
      `evidence.annotate(Image.new('RGB', (1600, 900), (248, 249, 250)), [(0, 103, 1600, 140)], 'Step 1: a full-width table', r'${out}')`,
      `r = Image.open(r'${out}').convert('RGB')`,
      "red = lambda x, y: r.getpixel((x, y)) == (255, 0, 0)",
      "print(int(any(red(x, 170) for x in range(0, 8))), int(any(red(x, 170) for x in range(1592, 1600))), int(any(red(800, y) for y in range(98, 112))), int(any(red(800, y) for y in range(236, 252))))",
    ].join("\n");
    const run = spawnSync(python, ["-c", code], { cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: temp() }, encoding: "utf8" });
    expect(run.status, run.stderr).toBe(0);
    expect(run.stdout.trim()).toBe("1 1 1 1");       // left, right, top, bottom
  });
});

describe("the test kit is used only when the person names it", () => {
  it.each(["/system-test run id=1", "  /system-test", "/SYSTEM-TEST plan select=all"])("%j is a /system-test request", (text) => {
    expect(SYSTEM_TEST_COMMAND.test(text)).toBe(true);
  });

  it.each(["test 3 case này cho tôi với", "chạy testcase ID 1", "please test this", "/system-testing x", "/opencode test it", "tính tổng cột D"])("%j is left alone", (text) => {
    expect(SYSTEM_TEST_COMMAND.test(text)).toBe(false);
  });
});

describe("kit: a relative QA_OUT", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("becomes a complete path inside the autotest folder, so the pictures it writes can be found from anywhere", () => {
    const home = temp();
    fs.writeFileSync(path.join(home, "qa.config.json"), "{}");
    const run = spawnSync(python, ["-c", "from qakit.core.config import cfg; print(cfg.out)"], {
      cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: home, QA_OUT: "test_execution_output/adhoc/dev_9" }, encoding: "utf8" });
    expect(run.status, run.stderr).toBe(0);
    expect(path.isAbsolute(run.stdout.trim())).toBe(true);
    expect(path.normalize(run.stdout.trim()).toLowerCase()).toBe(path.join(fs.realpathSync.native(home), "test_execution_output", "adhoc", "dev_9").toLowerCase());
  });
});

describe("kit: written for Python 3.9 and later (the person's computer is not this one)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("has no f-string whose expression holds a backslash or its own quote, which only Python 3.12 accepts (it broke every command on 3.11)", () => {
    const code = [
      "import ast, pathlib",
      "bad = []",
      "for f in sorted(pathlib.Path('qakit').rglob('*.py')):",
      "    src = f.read_text(encoding='utf-8')",
      "    for js in [n for n in ast.walk(ast.parse(src)) if isinstance(n, ast.JoinedStr)]:",
      "        whole = ast.get_source_segment(src, js) or ''",
      "        quote = whole.lstrip('rRfFbBuU')[:1]",
      "        triple = whole.lstrip('rRfFbBuU')[:3] in ('\"\"\"', \"'''\")",
      "        for fv in [n for n in ast.walk(js) if isinstance(n, ast.FormattedValue)]:",
      "            seg = ast.get_source_segment(src, fv.value) or ''",
      "            if '\\\\' in seg or (quote and quote in seg and not triple):",
      "                bad.append(f'{f}:{js.lineno}: {seg[:50]}')",
      "print('\\n'.join(bad))",
    ].join("\n");
    const run = spawnSync(python, ["-c", code], { cwd: path.join("autotest-kit", "kit"), encoding: "utf8" });
    expect(run.status, run.stderr).toBe(0);
    expect(run.stdout.trim()).toBe("");
  });
});

describe("a run that is repaired and run again inside itself", () => {
  const ev = (events: Array<Record<string, unknown>>) => events.map(e => E(e));
  const planned = { kind: "run_planned", cases: [{ suite: "orders", id: "A", steps: ["x"] }, { suite: "orders", id: "B", steps: ["x"] }] };

  it("is running again after a repair, and the case shows only its latest attempt", () => {
    const state = reduceRun("r", ev([planned, { kind: "run_started" }, { kind: "case_started", id: "A" }, { kind: "step", id: "A", n: 1, total: 1, status: "NOT EXECUTED", caption: "runner error", images: ["/old.png"] },
      { kind: "case_finished", id: "A", status: "NOT EXECUTED" }, { kind: "run_paused", reason: "case A needs repair", remaining: ["B"] }, { kind: "run_finished" },
      { kind: "run_resumed", reason: "again" }, { kind: "case_started", id: "A" }, { kind: "step", id: "A", n: 1, total: 1, status: "PASS", caption: "ok", images: ["/new.png"] }]), Date.parse("2026-11-17T10:00:05"));
    expect(state.phase).toBe("running");
    expect(state.stopReason).toBe("");
    const a = state.cases.find(item => item.id === "A")!;
    expect(a.status).toBe("RUNNING");
    expect(a.steps.map(step => step.status)).toEqual(["PASS"]);
    expect(a.steps[0]!.images).toEqual(["/new.png"]);
    expect(state.cases.find(item => item.id === "B")!.status).toBe("PLANNED");
  });

  it("is finished when the repaired run closes, with every case counted once", () => {
    const state = reduceRun("r", ev([planned, { kind: "run_started" }, { kind: "case_finished", id: "A", status: "NOT EXECUTED" }, { kind: "run_paused", reason: "case A needs repair" },
      { kind: "run_resumed" }, { kind: "case_started", id: "A" }, { kind: "case_finished", id: "A", status: "PASS" }, { kind: "case_started", id: "B" }, { kind: "case_finished", id: "B", status: "PASS" }, { kind: "run_finished" }]));
    expect(state.phase).toBe("finished");
    expect(state.counts).toEqual({ PASS: 2 });
  });

  it("keeps each step's own pictures and the pictures of a step still in progress apart", () => {
    const state = reduceRun("r", ev([{ kind: "case_started", id: "A" }, { kind: "shot", id: "A", n: 1, image: "/s1a.png" }, { kind: "step", id: "A", n: 1, total: 2, status: "PASS", caption: "one", images: ["/s1a.png", "/s1b.png"] },
      { kind: "shot", id: "A", n: 2, image: "/s2a.png" }]), Date.parse("2026-11-17T10:00:03"));
    const a = state.cases[0]!;
    expect(a.steps[0]!.images).toEqual(["/s1a.png", "/s1b.png"]);
    expect(a.liveImages).toEqual(["/s2a.png"]);
    expect(a.images).toEqual(["/s1a.png", "/s1b.png", "/s2a.png"]);
  });
});

describe("kit: watch decides from the state the log reaches (Python)", () => {
  const python = process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");
  const available = (() => { try { return spawnSync(python, ["--version"]).status === 0; } catch { return false; } })();

  it.skipIf(!available)("returns 10 while paused, ignores an old pause once the run resumed, and returns 0 when it finished", () => {
    const home = temp();
    fs.writeFileSync(path.join(home, "qa.config.json"), "{}");
    const dir = path.join(home, "test_execution_output", "runs", "r1");
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, "events.jsonl");
    const watch = (timeout = 1) => spawnSync(python, ["-m", "qakit.tools.watch", "--run", "r1", "--timeout", String(timeout)], { cwd: path.join("autotest-kit", "kit"), env: { ...process.env, QA_HOME: home }, encoding: "utf8" });
    fs.writeFileSync(file, line({ kind: "run_started" }) + line({ kind: "run_paused", reason: "case A needs repair" }));
    expect(watch().status).toBe(10);
    fs.appendFileSync(file, line({ kind: "run_resumed" }));
    expect(watch(2).status).toBe(11);                                     // resumed and still going: not stopped by the old pause
    fs.appendFileSync(file, line({ kind: "run_finished" }));
    expect(watch().status).toBe(0);
  });
});
