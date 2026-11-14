import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
// @ts-expect-error plain ESM module without type declarations
import { certificateInfo, ensureCertificate, generatePair, RENEW_BEFORE_DAYS } from "../bin/certificate.mjs";

const DAY = 86_400_000;
const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function tempDir() { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "op-cert-")); dirs.push(dir); return dir; }
function fakeTrust(accept = true) {
  const calls = { added: [] as string[], removed: [] as string[] };
  return { calls, trust: { add: (file: string) => { calls.added.push(certificateInfo(file).thumbprint); return accept; }, remove: (thumb: string) => { calls.removed.push(thumb); return true; } } };
}
async function install(dir: string, issuedAt: number) {
  const pair = await generatePair(issuedAt);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, "localhost.crt"), pair.cert);
  fs.writeFileSync(path.join(dir, "localhost.key"), pair.key);
  return certificateInfo(path.join(dir, "localhost.crt")).thumbprint as string;
}

describe("localhost certificate renewal", () => {
  it("keeps a certificate that is far from expiry and creates one when there is none", async () => {
    const dir = tempDir(), { trust, calls } = fakeTrust();
    const created = await ensureCertificate({ dir, trust });
    expect(created.status).toBe("created");
    expect(await ensureCertificate({ dir, trust })).toMatchObject({ status: "kept", thumbprint: created.thumbprint });
    expect(calls.added).toEqual([]); // setup trusts a new certificate itself
  });

  it("renews a certificate that expires within the renewal window, trusts the new one and untrusts the old one", async () => {
    const dir = tempDir(), { trust, calls } = fakeTrust();
    const old = await install(dir, Date.now());
    const now = Date.now() + (365 - RENEW_BEFORE_DAYS + 1) * DAY;
    const result = await ensureCertificate({ dir, trust, now });
    expect(result).toMatchObject({ status: "renewed", previous: old });
    expect(result.thumbprint).not.toBe(old);
    expect(calls.added).toEqual([result.thumbprint]);
    expect(calls.removed).toEqual([old]);
    expect(certificateInfo(path.join(dir, "localhost.crt")).validTo).toBeGreaterThan(now + 300 * DAY);
    expect(fs.readdirSync(dir).sort()).toEqual(["localhost.crt", "localhost.key"]);
  });

  it("keeps the working certificate when Windows refuses to trust the new one", async () => {
    const dir = tempDir(), { trust } = fakeTrust(false);
    const old = await install(dir, Date.now());
    await expect(ensureCertificate({ dir, trust, now: Date.now() + 340 * DAY })).rejects.toThrow(/CERTIFICATE_TRUST_FAILED/);
    expect(certificateInfo(path.join(dir, "localhost.crt")).thumbprint).toBe(old);
  });

  it("hands the renewed pair to the gateway that shared the old one and leaves the old trust for it", async () => {
    const dir = tempDir(), gateway = tempDir(), { trust, calls } = fakeTrust();
    const old = await install(dir, Date.now());
    fs.copyFileSync(path.join(dir, "localhost.crt"), path.join(gateway, "localhost.crt"));
    fs.copyFileSync(path.join(dir, "localhost.key"), path.join(gateway, "localhost.key"));
    const result = await ensureCertificate({ dir, sharedDir: gateway, trust, now: Date.now() + 340 * DAY });
    expect(certificateInfo(path.join(gateway, "localhost.crt")).thumbprint).toBe(result.thumbprint);
    expect(calls.removed).not.toContain(old);
  });
});
