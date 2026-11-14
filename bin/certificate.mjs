import { spawnSync } from "node:child_process";
import { X509Certificate } from "node:crypto";
import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";

/** The localhost certificate is replaced this long before it expires, so a whole fleet installed together never expires together. */
export const RENEW_BEFORE_DAYS = 60;
const DAY_MS = 86_400_000;
const VALID_DAYS = 365;

/** Thumbprint (SHA-1, as Windows shows it) and expiry of a certificate file; null when it is missing or unreadable. */
export function certificateInfo(file) {
  try {
    const certificate = new X509Certificate(readFileSync(file));
    return { thumbprint: certificate.fingerprint.replace(/:/g, "").toUpperCase(), validTo: new Date(certificate.validTo).getTime() };
  } catch { return null; }
}

export const needsRenewal = (info, now = Date.now()) => !info || info.validTo - now < RENEW_BEFORE_DAYS * DAY_MS;

/** A leaf certificate for localhost only (never a CA), so a leaked key cannot vouch for any other name. */
export async function generatePair(now = Date.now()) {
  const { default: selfsigned } = await import("selfsigned");
  const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }], {
    algorithm: "sha256",
    notAfterDate: new Date(now + VALID_DAYS * DAY_MS),
    extensions: [
      { name: "basicConstraints", cA: false },
      { name: "keyUsage", digitalSignature: true, keyEncipherment: true },
      { name: "extKeyUsage", serverAuth: true },
      { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }, { type: 7, ip: "::1" }] },
    ],
  });
  return { cert: pems.cert, key: pems.private };
}

/** The current Windows user's trusted roots (CurrentUser\Root: no administrator rights, other users unaffected). */
export const windowsTrust = {
  add: file => spawnSync("certutil.exe", ["-addstore", "-user", "Root", file], { stdio: "ignore", windowsHide: true }).status === 0,
  remove: thumbprint => spawnSync("certutil.exe", ["-user", "-delstore", "Root", thumbprint], { stdio: "ignore", windowsHide: true }).status === 0,
};

function writePair(dir, pair) {
  mkdirSync(dir, { recursive: true });
  for (const [name, text] of [["localhost.crt", pair.cert], ["localhost.key", pair.key]]) {
    const file = path.join(dir, name);
    writeFileSync(file + ".new", text, { mode: 0o600 });
    renameSync(file + ".new", file);
  }
}

/**
 * Makes `dir` hold a localhost certificate that stays valid for more than RENEW_BEFORE_DAYS.
 * - "kept": the current one is fine.
 * - "created": there was none; it is copied from `sharedDir` (the Agent Gateway, so the user trusts one certificate) or
 *   generated. Setup trusts it.
 * - "renewed": the current one expires soon or cannot be read. The new one is trusted before it replaces the old one, so
 *   a refused trust prompt leaves the working certificate in place. `sharedDir` gets the new pair when it held the old
 *   one (the gateway picks it up when it next starts); the old certificate's trust is removed only when nothing uses it.
 */
export async function ensureCertificate({ dir, sharedDir, trust = windowsTrust, generate = generatePair, now = Date.now() }) {
  const cert = path.join(dir, "localhost.crt"), key = path.join(dir, "localhost.key");
  const current = existsSync(key) ? certificateInfo(cert) : null;
  if (current && !needsRenewal(current, now)) return { status: "kept", thumbprint: current.thumbprint };
  const sharedCert = sharedDir && path.join(sharedDir, "localhost.crt"), sharedKey = sharedDir && path.join(sharedDir, "localhost.key");
  const shared = sharedDir && existsSync(sharedKey) ? certificateInfo(sharedCert) : null;
  if (!current && !existsSync(cert) && !existsSync(key)) {
    if (shared && !needsRenewal(shared, now)) {
      mkdirSync(dir, { recursive: true });
      copyFileSync(sharedCert, cert); copyFileSync(sharedKey, key);
      return { status: "created", thumbprint: shared.thumbprint, reused: true };
    }
    writePair(dir, await generate(now));
    return { status: "created", thumbprint: certificateInfo(cert).thumbprint };
  }
  // Renewal: a usable newer pair from the gateway first, else a new one.
  const pair = shared && !needsRenewal(shared, now) && shared.thumbprint !== current?.thumbprint
    ? { cert: readFileSync(sharedCert, "utf8"), key: readFileSync(sharedKey, "utf8") } : await generate(now);
  const staged = path.join(dir, "localhost.renew.crt");
  writeFileSync(staged, pair.cert, { mode: 0o600 });
  const trusted = trust.add(staged);
  rmSync(staged, { force: true });
  if (!trusted) throw new Error("CERTIFICATE_TRUST_FAILED: Windows did not trust the renewed localhost certificate; the current one is kept.");
  writePair(dir, pair);
  const next = certificateInfo(cert);
  let gatewayUsesOld = false;
  if (current && shared?.thumbprint === current.thumbprint) {
    writePair(sharedDir, pair); // the running gateway keeps the old one in memory until it restarts
    gatewayUsesOld = true;
  }
  if (current && current.thumbprint !== next.thumbprint && !gatewayUsesOld) trust.remove(current.thumbprint);
  return { status: "renewed", thumbprint: next.thumbprint, previous: current?.thumbprint ?? null };
}
