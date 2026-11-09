#!/usr/bin/env node
// `excel-partner` is the old name of `office-partner`. See migrate.mjs.
import { forward } from "./migrate.mjs";

process.exitCode = forward(process.argv.slice(2));
