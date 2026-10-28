// Excel Partner QA harness (managed file). Environment variables are set by `excel-partner qa run`.
import path from 'node:path';
import { defineConfig } from '@playwright/test';
import { launchArgs } from './_harness/launch.js';

const runDir = process.env.QA_RUN_DIR || path.resolve('results', 'adhoc');

export default defineConfig({
  testDir: './tests',
  testMatch: '**/*.spec.js',
  outputDir: path.join(runDir, 'artifacts'),
  timeout: Number(process.env.QA_TIMEOUT_MS || 60000),
  expect: { timeout: 10000 },
  workers: Number(process.env.QA_WORKERS || 1),
  retries: Number(process.env.QA_RETRIES || 0),
  fullyParallel: false,
  forbidOnly: true,
  reporter: [['./_harness/reporter.js'], ['line']],
  use: {
    headless: process.env.QA_HEADED !== '1',
    screenshot: 'only-on-failure',
    trace: process.env.QA_TRACE === '1' ? 'retain-on-failure' : 'off',
    video: 'off',
    launchOptions: { args: launchArgs() },
  },
});
