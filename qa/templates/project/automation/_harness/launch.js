// Excel Partner QA harness (managed file: changes are detected and blocked until `excel-partner qa upgrade`).
// Builds browser launch options so the browser can only reach the allowed origins.

export function allowedOrigins() {
  return (process.env.QA_ALLOWED_ORIGINS || '').split(',').map(value => value.trim()).filter(Boolean);
}

export function isAllowedUrl(url, origins = allowedOrigins()) {
  let parsed;
  try { parsed = new URL(url); } catch { return false; }
  if (['data:', 'blob:', 'about:'].includes(parsed.protocol)) return true;
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:' && parsed.protocol !== 'ws:' && parsed.protocol !== 'wss:') return false;
  const origin = parsed.protocol.replace(/^ws/, 'http') + '//' + parsed.host;
  return origins.includes(origin);
}

/**
 * Chromium switches that make every host except the allowed ones unreachable, whatever the page script does:
 * - requests go to a dead proxy unless the destination is in the bypass list (this also covers IP addresses,
 *   which the resolver rules below do not), and
 * - DNS answers "not found" for every other name.
 */
export function launchArgs(origins = allowedOrigins()) {
  const bypass = [], hosts = [];
  for (const origin of origins) {
    let url;
    try { url = new URL(origin); } catch { continue; }
    const host = url.hostname.includes(':') && !url.hostname.startsWith('[') ? `[${url.hostname}]` : url.hostname;
    bypass.push(url.port ? `${host}:${url.port}` : `${url.protocol}//${host}`);
    hosts.push(host.replace(/^\[|\]$/g, ''));
  }
  const rules = ['MAP * ~NOTFOUND', ...[...new Set(hosts)].map(host => `EXCLUDE ${host}`)];
  return [
    '--proxy-server=http://127.0.0.1:9',
    `--proxy-bypass-list=<-loopback>;${[...new Set(bypass)].join(';')}`,
    `--host-resolver-rules=${rules.join(' , ')}`,
    '--disable-background-networking', '--no-default-browser-check',
  ];
}
