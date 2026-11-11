/**
 * Secret scrubbing. Every path that prints — logs, errors, --dry-run output,
 * `auth status` — routes through here. Nothing else is trusted to remember.
 */

/** Keys whose values are always secret, regardless of shape. */
const SECRET_KEYS = new Set([
  'access_token',
  'refresh_token',
  'id_token',
  'device_code',
  'client_secret',
  'accesstoken',
  'refreshtoken',
  'authorization',
  'code_verifier',
  // base64-encoded file/attachment content — never print in dry-run output
  'contentbytes',
]);

/** JWTs and long opaque bearer blobs. */
const JWT_RE = /\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g;
const BEARER_RE = /\b(Bearer\s+)[A-Za-z0-9._~+/-]{20,}=*/gi;
/** Microsoft refresh tokens: long base64url-ish runs, often prefixed 0.A / M.C / 1.A */
const OPAQUE_TOKEN_RE = /\b[0-9A-Za-z]\.[A-Za-z0-9_-]{40,}/g;

export const MASK = '[REDACTED]';

/** Mask a known-secret string, keeping a short tail for correlation. */
export function maskSecret(value: string): string {
  if (!value) return MASK;
  if (value.length <= 8) return MASK;
  return `${MASK}(****${value.slice(-4)})`;
}

/** Scrub secrets out of free text (error bodies, HTTP traces). */
export function redactText(text: string): string {
  if (!text) return text;
  return text
    .replace(JWT_RE, MASK)
    .replace(BEARER_RE, (_m, p1: string) => `${p1}${MASK}`)
    .replace(OPAQUE_TOKEN_RE, MASK);
}

/**
 * Deep-clone a value with every secret-keyed field masked.
 * Used before any JSON.stringify that could reach a terminal or log.
 */
export function redactObject<T>(input: T): T {
  return walk(input, 0) as T;
}

function walk(value: unknown, depth: number): unknown {
  // Past the cap this must mask, not pass through: returning `value` here let
  // a secret nested one level deeper than the cap print raw, defeating the
  // one guarantee this module exists to make.
  if (depth > 12) return MASK;

  if (typeof value === 'string') return redactText(value);
  if (value === null || typeof value !== 'object') return value;

  if (Array.isArray(value)) {
    return value.map((item) => walk(item, depth + 1));
  }

  const out: Record<string, unknown> = {};
  for (const [key, val] of Object.entries(value as Record<string, unknown>)) {
    if (SECRET_KEYS.has(key.toLowerCase())) {
      out[key] = typeof val === 'string' ? maskSecret(val) : MASK;
    } else {
      out[key] = walk(val, depth + 1);
    }
  }
  return out;
}

/** Redact an outbound header map for --dry-run display. */
export function redactHeaders(
  headers: Record<string, string>,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, val] of Object.entries(headers)) {
    out[key] = SECRET_KEYS.has(key.toLowerCase()) ? MASK : redactText(val);
  }
  return out;
}
