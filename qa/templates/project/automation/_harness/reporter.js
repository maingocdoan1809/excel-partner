// Excel Partner QA harness (managed file).
// Writes <run dir>/results.json: one record per test case id, with credentials removed from error text.
import fs from 'node:fs';
import path from 'node:path';
import { redact } from './redact.js';

// eslint-disable-next-line no-control-regex
const ANSI = /\u001b\[[0-9;]*m/g;
const ID = /^\[([^\]]+)\]/;

/** Redact every string inside a parsed value. */
function redactDeep(value) {
  if (typeof value === 'string') return redact(value);
  if (Array.isArray(value)) return value.map(redactDeep);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, redactDeep(item)]));
  return value;
}

/** Title of the innermost step that failed, such as "locator.click". */
function failedStep(steps) {
  for (const step of steps || []) {
    if (step.error) return failedStep(step.steps) || step.title;
  }
  return undefined;
}

export default class QaReporter {
  constructor() {
    this.projectDir = process.env.QA_PROJECT_DIR || process.cwd();
    this.runDir = process.env.QA_RUN_DIR;
    this.records = new Map();
    this.startedAt = new Date().toISOString();
  }

  onBegin() { this.startedAt = new Date().toISOString(); }

  onTestEnd(test, result) {
    const match = ID.exec(test.title);
    const tcId = match ? match[1] : test.title;
    const outcome = test.outcome();
    const status = outcome === 'expected' ? 'passed' : outcome === 'flaky' ? 'flaky' : outcome === 'skipped' ? 'skipped' : result.status === 'timedOut' ? 'timedOut' : 'failed';
    const evidence = [];
    for (const attachment of result.attachments) {
      if (attachment.path) evidence.push(path.relative(this.projectDir, attachment.path).split(path.sep).join('/'));
    }
    let pageState;
    const stateFile = result.attachments.find(attachment => attachment.name === 'page-state' && attachment.body);
    if (stateFile && stateFile.body.length <= 12000) {
      try { pageState = redactDeep(JSON.parse(stateFile.body.toString('utf8'))); } catch { /* unreadable: leave it out */ }
    }
    const location = result.error && result.error.location;
    const failedAt = location ? { file: path.basename(location.file), line: location.line, column: location.column } : undefined;
    const snippet = result.error && result.error.snippet ? redact(String(result.error.snippet).replace(ANSI, '')).slice(0, 800) : undefined;
    const previous = this.records.get(tcId);
    const error = result.error ? redact(String(result.error.message || result.error).replace(ANSI, '')).slice(0, 2000) : undefined;
    this.records.set(tcId, {
      tcId, title: test.title.replace(ID, '').trim(), status,
      durationMs: (previous?.durationMs ?? 0) + result.duration,
      retries: result.retry, error: status === 'passed' ? previous?.error : (error ?? previous?.error),
      evidence: [...new Set([...(previous?.evidence ?? []), ...evidence])],
      // What the last failing attempt looked like. A later passing attempt keeps it, so a flaky test can still be investigated.
      pageState: status === 'passed' ? previous?.pageState : (pageState ?? previous?.pageState),
      failedAt: status === 'passed' ? previous?.failedAt : (failedAt ?? previous?.failedAt),
      failedStep: status === 'passed' ? previous?.failedStep : (failedStep(result.steps) ?? previous?.failedStep),
      snippet: status === 'passed' ? previous?.snippet : (snippet ?? previous?.snippet),
    });
  }

  onError(error) { this.fatal = redact(String(error.message || error).replace(ANSI, '')).slice(0, 2000); }

  onEnd(result) {
    if (!this.runDir) return;
    const tests = [...this.records.values()];
    const count = status => tests.filter(test => test.status === status).length;
    const output = {
      runId: path.basename(this.runDir), startedAt: this.startedAt, finishedAt: new Date().toISOString(),
      baseUrl: process.env.QA_BASE_URL, environment: process.env.QA_ENVIRONMENT, overall: result.status,
      summary: { total: tests.length, passed: count('passed'), failed: count('failed') + count('timedOut'), flaky: count('flaky'), skipped: count('skipped') },
      ...(this.fatal ? { fatal: this.fatal } : {}),
      tests,
    };
    fs.mkdirSync(this.runDir, { recursive: true });
    fs.writeFileSync(path.join(this.runDir, 'results.json'), JSON.stringify(output, null, 2));
  }
}
