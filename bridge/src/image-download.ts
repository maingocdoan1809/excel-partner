import dns from "node:dns";
import http from "node:http";
import https from "node:https";
import net from "node:net";

export class ImageDownloadError extends Error {}

export interface DownloadOptions {
  /** The download is cut off, and refused, past this many bytes. */
  maxBytes: number;
  timeoutMs?: number;
  /** Aborts the download when the request that asked for it ends. */
  signal?: AbortSignal;
  /** Tests only: reach loopback and private addresses. Off in the product. */
  allowPrivateNetwork?: boolean;
  maxRedirects?: number;
}

export interface Downloaded {
  bytes: Buffer;
  /** The address the bytes came from after redirects. */
  finalUrl: string;
}

/** The download the runtime uses; a different one can be passed in (tests). */
export type ImageDownloader = (url: string, options: DownloadOptions) => Promise<Downloaded>;

const DEFAULT_TIMEOUT_MS = 20_000;
const DEFAULT_REDIRECTS = 4;

function ipv4Private(address: string): boolean {
  const [a, b, c] = address.split(".").map(Number) as [number, number, number];
  return a === 0 || a === 10 || a === 127 || a >= 224
    || (a === 100 && b >= 64 && b <= 127)
    || (a === 169 && b === 254)
    || (a === 172 && b >= 16 && b <= 31)
    || (a === 192 && b === 0 && c === 0)
    || (a === 192 && b === 168)
    || (a === 198 && (b === 18 || b === 19));
}

/** Loopback, private, link-local, multicast and otherwise non-public addresses: what a page on the internet should never make this computer open. */
export function isPrivateAddress(address: string): boolean {
  const bare = address.replace(/^\[|\]$/g, "").replace(/%.*$/, "").toLowerCase();
  if (net.isIPv4(bare)) return ipv4Private(bare);
  if (!net.isIPv6(bare)) return true;
  const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(bare);
  if (mapped) return ipv4Private(mapped[1]!);
  const mappedHex = /^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(bare);
  if (mappedHex) {
    const high = parseInt(mappedHex[1]!, 16);
    const low = parseInt(mappedHex[2]!, 16);
    return ipv4Private(`${high >> 8}.${high & 255}.${low >> 8}.${low & 255}`);
  }
  const first = parseInt(bare.split(":")[0] || "0", 16);
  return bare === "::" || bare === "::1" || (first & 0xfe00) === 0xfc00 || (first & 0xffc0) === 0xfe80 || (first & 0xff00) === 0xff00;
}

type LookupCallback = (error: NodeJS.ErrnoException | null, address: string | dns.LookupAddress[], family?: number) => void;

/** DNS lookup that refuses a name resolving to a non-public address. It runs for the connection itself, so a name cannot change its answer between check and use. */
function publicLookup(hostname: string, options: dns.LookupOptions, callback: LookupCallback): void {
  dns.lookup(hostname, { ...options, all: true }, (error, addresses) => {
    if (error) return callback(error, "");
    const list = addresses as dns.LookupAddress[];
    if (list.length === 0 || list.some((entry) => isPrivateAddress(entry.address))) {
      return callback(Object.assign(new Error(`${hostname} is not a public address`), { code: "EACCES" }), "");
    }
    if (options.all) return callback(null, list);
    callback(null, list[0]!.address, list[0]!.family);
  });
}

function parseUrl(text: string, options: DownloadOptions): URL {
  let url: URL;
  try { url = new URL(text); } catch { throw new ImageDownloadError(`'${text.slice(0, 200)}' is not a valid address`); }
  if (url.protocol !== "https:" && url.protocol !== "http:") throw new ImageDownloadError("only http:// and https:// addresses can be downloaded");
  if (url.username || url.password) throw new ImageDownloadError("addresses that carry a user name or password are not accepted");
  if (!options.allowPrivateNetwork && net.isIP(url.hostname.replace(/^\[|\]$/g, "")) !== 0 && isPrivateAddress(url.hostname)) {
    throw new ImageDownloadError(`${url.hostname} is not a public address; only pictures on the internet can be downloaded`);
  }
  return url;
}

interface Hop { status: number; location?: string; bytes?: Buffer }

function fetchOnce(url: URL, options: DownloadOptions, signal: AbortSignal): Promise<Hop> {
  return new Promise((resolve, reject) => {
    const client = url.protocol === "https:" ? https : http;
    const request = client.request(url, {
      method: "GET",
      headers: { "user-agent": "OfficePartner/1.0 (picture download)", accept: "image/png,image/jpeg,image/gif,image/*;q=0.8", "accept-encoding": "identity" },
      ...(options.allowPrivateNetwork ? {} : { lookup: publicLookup as unknown as net.LookupFunction }),
      signal,
    }, (response) => {
      const status = response.statusCode ?? 0;
      if (status >= 300 && status < 400 && response.headers.location) {
        response.resume();
        return resolve({ status, location: response.headers.location });
      }
      if (status < 200 || status >= 300) {
        response.resume();
        return reject(new ImageDownloadError(`the server answered HTTP ${status}`));
      }
      const declared = Number(response.headers["content-length"]);
      if (Number.isFinite(declared) && declared > options.maxBytes) {
        response.destroy();
        return reject(new ImageDownloadError(`the picture is ${declared} bytes, over the ${options.maxBytes} byte limit`));
      }
      const chunks: Buffer[] = [];
      let total = 0;
      response.on("data", (chunk: Buffer) => {
        total += chunk.length;
        if (total > options.maxBytes) {
          response.destroy();
          reject(new ImageDownloadError(`the picture is larger than the ${options.maxBytes} byte limit`));
          return;
        }
        chunks.push(chunk);
      });
      response.on("end", () => resolve({ status, bytes: Buffer.concat(chunks) }));
      response.on("error", reject);
    });
    request.on("error", (error: NodeJS.ErrnoException) => {
      if (signal.aborted) return reject(new ImageDownloadError("the download was cancelled or took too long"));
      reject(new ImageDownloadError(error.code === "EACCES" ? `${url.hostname} is not a public address; only pictures on the internet can be downloaded` : `could not download the picture: ${error.message}`));
    });
    request.end();
  });
}

/** Download a picture from a public web address: http(s) only, no addresses on this computer or its network, a size limit, a time limit, a few redirects. */
export const downloadImage: ImageDownloader = async (address, options) => {
  const redirects = options.maxRedirects ?? DEFAULT_REDIRECTS;
  const signal = AbortSignal.any([AbortSignal.timeout(options.timeoutMs ?? DEFAULT_TIMEOUT_MS), ...(options.signal ? [options.signal] : [])]);
  let url = parseUrl(address, options);
  for (let hop = 0; hop <= redirects; hop++) {
    const result = await fetchOnce(url, options, signal);
    if (result.bytes) return { bytes: result.bytes, finalUrl: url.href };
    url = parseUrl(new URL(result.location!, url).href, options);
  }
  throw new ImageDownloadError(`the address redirects more than ${redirects} times`);
};
