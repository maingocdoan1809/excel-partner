import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { Transform } from "node:stream";
import { pipeline } from "node:stream/promises";

/** Manifest-owned native releases. Downloaded bytes must match the release pin before anything is extracted or run. */
export function validDownload(item) {
  return item && /^https:\/\/[^\s]+$/.test(item.url) && /^[a-f0-9]{128}$/.test(item.sha512)
    && (!item.archiveEntry || /^[a-zA-Z0-9_.-]+$/.test(item.archiveEntry) && ![".", ".."].includes(item.archiveEntry));
}
export async function installBinary(directory, item, { run, env, fetchRelease = fetch } = {}) {
  const release = item.download?.[`${process.platform}-${process.arch}`];
  if (!release) throw new Error("RUNTIME_PLATFORM_UNSUPPORTED");
  if (!validDownload(release)) throw new Error("RUNTIME_MANIFEST_INVALID");
  const payload = path.join(directory, ".download");
  const target = path.join(directory, item.bin + (process.platform === "win32" ? ".exe" : ""));
  const response = await fetchRelease(release.url, { signal: AbortSignal.timeout(10 * 60_000) });
  if (!response.ok || !response.body) throw new Error("INSTALL_FAILED");
  const hash = crypto.createHash("sha512");
  let size = 0;
  try {
    await pipeline(response.body, new Transform({ transform(chunk, _encoding, callback) {
      size += chunk.length;
      if (size > 512 * 1024 * 1024) return callback(new Error("INSTALL_FAILED"));
      hash.update(chunk); callback(null, chunk);
    } }), fs.createWriteStream(payload, { flags: "wx", mode: 0o600 }));
    if (hash.digest("hex") !== release.sha512) throw new Error("RUNTIME_INTEGRITY_FAILED");
    if (release.archiveEntry) {
      const result = await run("tar", ["-xzf", payload, "-C", directory, "--", release.archiveEntry], { env, timeoutMs: 60_000 });
      if (result.code !== 0) throw new Error("INSTALL_FAILED");
      const extracted = path.join(directory, release.archiveEntry);
      if (!fs.lstatSync(extracted).isFile()) throw new Error("CLI_ENTRYPOINT_INVALID");
      fs.renameSync(extracted, target);
    } else fs.renameSync(payload, target);
    fs.chmodSync(target, 0o700);
  } finally { fs.rmSync(payload, { force: true }); }
}
