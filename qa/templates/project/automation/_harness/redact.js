// Excel Partner QA harness (managed file).
// Removes values that look like credentials from text that is stored or shown.

export function redact(text) {
  return String(text)
    .replace(/(authorization|proxy-authorization)(["']?\s*[:=]\s*["']?)(bearer|basic|token)?\s*[^\s"',;]+/gi, '$1$2[redacted]')
    .replace(/((?:set-)?cookie["']?\s*[:=]\s*["']?)[^\n"']+/gi, '$1[redacted]')
    .replace(/\b(session_id|sessionid|access_token|refresh_token|id_token|csrf_token|api[_-]?key|secret|password|passwd|pwd)(["']?\s*[:=]\s*["']?)[^\s"',;&]+/gi, '$1$2[redacted]')
    .replace(/\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/g, '[redacted-jwt]')
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]{12,}/gi, 'Bearer [redacted]');
}
