/**
 * Parsing for `outlook auth login --scopes <comma list>`.
 *
 * A custom scope set lets one job (an unattended mail-moving agent, say) hold
 * a token that is narrower than the default grant, e.g. only `Mail.ReadWrite`.
 * The exact set requested is persisted with the token (`grantScopes`) so that
 * `auth refresh` asks Azure for the SAME set again instead of DEFAULT_SCOPES —
 * otherwise the first refresh would quietly widen the grant back to the full
 * scope list.
 */

import { UsageError } from './errors.js';

/** Graph scope tokens: `Mail.ReadWrite`, `offline_access`, or a full resource URI. */
const SCOPE_RE = /^[A-Za-z0-9._:/-]+$/;

const OFFLINE_ACCESS = 'offline_access';

/**
 * Split a comma (or whitespace) separated scope list, validate each token,
 * drop case-insensitive duplicates, and make sure `offline_access` is present
 * — without it Azure issues no refresh token and the session dies in an hour.
 */
export function parseScopeList(raw: string): string[] {
  const tokens = raw
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);

  if (tokens.length === 0) {
    throw new UsageError('--scopes needs at least one scope, e.g. --scopes Mail.ReadWrite');
  }

  const seen = new Set<string>();
  const scopes: string[] = [];
  for (const token of tokens) {
    if (!SCOPE_RE.test(token)) {
      throw new UsageError(`--scopes: "${token}" is not a valid scope name.`);
    }
    const key = token.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    scopes.push(token);
  }

  if (!seen.has(OFFLINE_ACCESS)) scopes.unshift(OFFLINE_ACCESS);
  return scopes;
}

/** Runtime guard for the persisted `grantScopes` field. */
export function isScopeArray(value: unknown): value is string[] {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.every((s) => typeof s === 'string' && SCOPE_RE.test(s))
  );
}
