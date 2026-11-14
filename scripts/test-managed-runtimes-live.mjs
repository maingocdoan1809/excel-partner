/** Opt-in real package installs and native handshakes, with isolated provider state. No model prompts. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import store from '../bin/runtime-store.cjs';
import { ensureBaseline, installRecovery } from '../bin/manage.mjs';
import { installRuntime, readManifest, runtimeInventory } from '../bin/runtime-manager.mjs';
import { runCommand } from '../bin/doctor-probes.mjs';

const providers = (process.env.LIVE_MANAGED_RUNTIMES ?? '').split(',').map(id => id.trim()).filter(Boolean);
if (!providers.length || providers.some(id => !store.ids.includes(id))) {
  throw new Error('Opt in with LIVE_MANAGED_RUNTIMES=copilot (or a comma-separated provider list).');
}
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'output', 'playwright', 'runtime-live');
const sources = ['bin', 'dist/providers/src'].flatMap(directory => fs.readdirSync(path.join(root, directory))
  .filter(file => /\.(?:mjs|cjs|js|json)$/.test(file)).sort().map(file => path.join(directory, file)));
const revision = store.digest(sources.map(file => file + ':' + store.fileDigest(path.join(root, file))).join('\n')).slice(0, 20);
const dataDir = path.join(output, 'data-' + revision), runtimeData = path.relative(root, dataDir);
fs.mkdirSync(output, { recursive: true });
console.log('Preparing isolated app snapshot. Existing provider installations and sign-ins are kept.');
installRecovery(root, dataDir);
const baseline = ensureBaseline(root, dataDir);
let prior = [];
try { const saved = JSON.parse(fs.readFileSync(path.join(output, 'results.json'), 'utf8')); if (saved.runtimeData === runtimeData) prior = saved.checks ?? []; } catch { /* first check */ }
const checks = prior.filter(check => !providers.includes(check.provider)), manifest = readManifest(root);
for (const provider of providers) {
  const operations = [];
  console.log(`Checking private ${provider} ${manifest.providers[provider].version}...`);
  const run = async (command, args, options) => {
    const result = await runCommand(command, args, options);
    if (result.code !== 0) {
      const npmCode = /\b(?:EBADENGINE|ETARGET|E404|ETIMEDOUT|ENOTFOUND|EAI_AGAIN|ECONNRESET|CERT_HAS_EXPIRED)\b/.exec(result.stderr ?? '')?.[0];
      operations.push({ exitCode: result.code, errorCode: result.errorCode ?? npmCode ?? null });
    }
    return result;
  };
  try {
    const profile = await installRuntime(root, provider, { dataDir, run });
    checks.push({ provider, version: profile.version, ok: true, nativePreflight: true, operations });
    console.log(`PASS ${provider}: exact CLI, integrity and native adapter preflight verified.`);
  } catch (error) {
    const code = String(error.code ?? error.message).match(/^[A-Z][A-Z0-9_]+$/)?.[0] ?? 'LIVE_RUNTIME_CHECK_FAILED';
    checks.push({ provider, version: manifest.providers[provider].version, ok: false, code, operations });
    console.log(`FAIL ${provider}: ${code}. Previous selection retained.`);
  }
  fs.writeFileSync(path.join(output, 'results.json'), JSON.stringify({ runtimeData, checks, providers: runtimeInventory(dataDir) }, null, 2));
}
if (checks.some(check => !check.ok)) process.exitCode = 1;
