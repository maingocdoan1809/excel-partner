// Every test runs against an empty data folder, so none of them can read (or change) the installed product's providers, sign-ins or history.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterAll } from "vitest";

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "office-partner-test-"));
process.env.OFFICE_PARTNER_DATA_DIR = dataDir;

afterAll(() => {
  try { fs.rmSync(dataDir, { recursive: true, force: true }); } catch { /* a leftover temp folder is harmless */ }
});
