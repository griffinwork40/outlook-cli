/**
 * Token acquisition, refresh, and status — the surface the CLI's commands
 * actually call. Wraps deviceCode.ts (initiate + poll) and tokenStore.ts
 * (disk cache) behind a single `Authenticator`.
 *
 * Refresh-token rotation (docs/00 §4) is load-bearing: Microsoft may return
 * a NEW refresh token on every refresh call, and the old one can be silently
 * invalidated once that happens. Every refresh here persists whatever
 * refresh_token comes back — never the one that was just spent.
 */

import { authorityBase } from './config.js';
import { NotAuthenticatedError, OutlookCliError, parseGraphError } from './errors.js';
import { requestDeviceCode, pollForToken } from './deviceCode.js';
import { postForm } from './identityFetch.js';
import type { TokenStore } from './tokenStore.js';
import type {
  DeviceCodeResponse,
  GrantMode,
  OutlookConfig,
  TokenResponse,
  TokenSet,
} from './types.js';
import { scopesForMode } from './types.js';

/** getAccessToken refreshes proactively once the token is this close to expiry. */
const REFRESH_SKEW_MS = 60_000;

export interface AuthStatus {
  signedIn: boolean;
  account?: string;
  expiresAt?: number;
  expired: boolean;
  scopes?: string;
  /** Absent on a pre-existing token issued before read-only mode existed. */
  mode?: GrantMode;
  /** Present only when the session was signed in with `--scopes`. */
  grantScopes?: string[];
}

export interface LoginOptions {
  /** Request the reduced scope set — cannot send, delete, or modify mail. */
  readOnly?: boolean;
  /**
   * Request exactly these scopes (already normalized by parseScopeList, so
   * `offline_access` is included). Mutually exclusive with readOnly.
   */
  scopes?: readonly string[];
}

export interface Authenticator {
  getAccessToken(): Promise<string>;
  login(
    onPrompt: (d: DeviceCodeResponse) => void,
    opts?: LoginOptions,
  ): Promise<TokenSet>;
  logout(): Promise<void>;
  refresh(): Promise<TokenSet>;
  status(): Promise<AuthStatus>;
}

export interface AuthenticatorDeps {
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  /** Injectable clock so expiry math is deterministic in tests. */
  now?: () => number;
}

export function createAuthenticator(
  cfg: OutlookConfig,
  store: TokenStore,
  deps: AuthenticatorDeps = {},
): Authenticator {
  const fetchImpl = deps.fetchImpl ?? fetch;
  const sleep = deps.sleep ?? defaultSleep;
  const now = deps.now ?? Date.now;

  async function refresh(): Promise<TokenSet> {
    const cached = await store.load();
    if (!cached) throw new NotAuthenticatedError();
    if (!cached.refreshToken) {
      throw new NotAuthenticatedError('cached session has no refresh token; run login again');
    }

    // Refresh with the scope set this grant was ORIGINALLY obtained under.
    // Priority (highest to lowest):
    //   1. cached.grantScopes — set by `--scopes` login; exact list, already
    //      includes offline_access.
    //   2. cached.scope — the scope string Azure echoed back at login time;
    //      used for tokens from older builds (e.g. a fork binary that added
    //      Contacts.ReadWrite before main did). offline_access is stripped by
    //      Azure on echo, so we re-add it here.  This prevents a binary
    //      upgrade from silently narrowing a grant that the user obtained under
    //      a wider scope set.
    //   3. scopesForMode(mode) — fallback for legacy tokens that pre-date
    //      cached.scope, or for tokens where Azure echoed an empty string.
    //
    // A `--readonly` session must NEVER widen: when mode is 'readonly' the
    // cached.scope will already be the read-only set (Azure echoes back only
    // what was granted), so rule 2 is safe — it cannot escalate.
    const mode = cached.mode ?? 'full';
    const scopes = cached.grantScopes ?? scopesFromCachedScope(cached.scope) ?? scopesForMode(mode);
    const tokenResponse = await exchangeRefreshToken(
      cfg,
      cached.refreshToken,
      scopes,
      fetchImpl,
    );
    const nextTokens = toTokenSet(tokenResponse, now(), cached);
    await store.save(nextTokens);
    return nextTokens;
  }

  return {
    async getAccessToken(): Promise<string> {
      const cached = await store.load();
      if (!cached) throw new NotAuthenticatedError();

      if (cached.expiresAt - now() < REFRESH_SKEW_MS) {
        const refreshed = await refresh();
        return refreshed.accessToken;
      }
      return cached.accessToken;
    },

    async login(
      onPrompt: (d: DeviceCodeResponse) => void,
      opts: LoginOptions = {},
    ): Promise<TokenSet> {
      const mode: GrantMode = opts.readOnly ? 'readonly' : 'full';
      const custom = opts.scopes !== undefined ? [...opts.scopes] : undefined;
      const deviceCode = await requestDeviceCode(
        cfg,
        custom ?? scopesForMode(mode),
        fetchImpl,
      );
      onPrompt(deviceCode);

      const tokenResponse = await pollForToken(cfg, deviceCode.device_code, {
        intervalSec: deviceCode.interval,
        expiresInSec: deviceCode.expires_in,
        fetchImpl,
        sleep,
      });

      const tokens: TokenSet = {
        ...toTokenSet(tokenResponse, now()),
        mode,
        ...(custom !== undefined ? { grantScopes: custom } : {}),
      };
      await store.save(tokens);
      return tokens;
    },

    async logout(): Promise<void> {
      await store.clear();
    },

    refresh,

    async status(): Promise<AuthStatus> {
      const cached = await store.load();
      if (!cached) return { signedIn: false, expired: false };

      return {
        signedIn: true,
        expired: now() >= cached.expiresAt,
        expiresAt: cached.expiresAt,
        ...(cached.account !== undefined ? { account: cached.account } : {}),
        ...(cached.scope !== undefined ? { scopes: cached.scope } : {}),
        ...(cached.mode !== undefined ? { mode: cached.mode } : {}),
        ...(cached.grantScopes !== undefined ? { grantScopes: cached.grantScopes } : {}),
      };
    },
  };
}

/** docs/00 §4: same /token endpoint, grant_type=refresh_token, no client_secret for a public client. */
async function exchangeRefreshToken(
  cfg: OutlookConfig,
  refreshToken: string,
  scopes: readonly string[],
  fetchImpl: typeof fetch,
): Promise<TokenResponse> {
  const url = `${authorityBase(cfg.tenantId)}/token`;
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: scopes.join(' '),
  });

  const res = await postForm(url, params, fetchImpl);
  const bodyText = await res.text().catch(() => '');

  if (!res.ok) {
    throw parseGraphError(res.status, bodyText);
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(bodyText);
  } catch {
    throw new OutlookCliError(
      'The token endpoint returned a 200 response with an unparseable body during refresh.',
    );
  }
  if (!isTokenResponse(parsed)) {
    throw new OutlookCliError(
      'The token endpoint returned a 200 response with an unrecognized body shape during refresh.',
    );
  }
  return parsed;
}

/**
 * Builds the TokenSet to persist. Rotation contract: `refresh_token` from
 * THIS response wins whenever present; we fall back to `previous`'s only
 * when the server omitted one on this call (still oscillates toward the
 * newest token Microsoft has issued, never re-sends a stale one).
 */
function toTokenSet(
  tokenResponse: TokenResponse,
  nowMs: number,
  previous?: TokenSet,
): TokenSet {
  const account = tokenResponse.id_token
    ? (decodeIdTokenAccount(tokenResponse.id_token) ?? previous?.account)
    : previous?.account;
  const refreshToken = tokenResponse.refresh_token ?? previous?.refreshToken;
  const scope = tokenResponse.scope ?? previous?.scope;
  const tokenType = tokenResponse.token_type ?? previous?.tokenType;
  // The grant mode is OURS, not Azure's — the token response never echoes it,
  // so it can only survive a refresh by being carried forward here. Dropping
  // it would silently re-escalate: the next refresh would read `mode ?? 'full'`
  // and ask for the full scope set, and guardMutation's read-only gate would
  // stop firing. A --readonly session must stay read-only across refreshes.
  const mode = previous?.mode;
  // Same for a custom scope set: drop it and the next refresh asks for
  // DEFAULT_SCOPES, widening a deliberately narrow grant.
  const grantScopes = previous?.grantScopes;

  return {
    accessToken: tokenResponse.access_token,
    expiresAt: nowMs + tokenResponse.expires_in * 1000,
    ...(refreshToken !== undefined ? { refreshToken } : {}),
    ...(scope !== undefined ? { scope } : {}),
    ...(tokenType !== undefined ? { tokenType } : {}),
    ...(account !== undefined ? { account } : {}),
    ...(mode !== undefined ? { mode } : {}),
    ...(grantScopes !== undefined ? { grantScopes } : {}),
  };
}

/**
 * Best-effort account resolution from an id_token's payload claims.
 * Signature is NEVER verified here — this is display metadata only, not an
 * authorization decision, and the token itself is discarded after this call
 * (TokenSet has no id_token field; nothing derived from it is ever logged).
 */
function decodeIdTokenAccount(idToken: string): string | undefined {
  try {
    const parts = idToken.split('.');
    if (parts.length !== 3) return undefined;
    const payloadSegment = parts[1];
    if (!payloadSegment) return undefined;

    const json = Buffer.from(payloadSegment, 'base64url').toString('utf8');
    const claims: unknown = JSON.parse(json);
    if (typeof claims !== 'object' || claims === null) return undefined;

    const c = claims as Record<string, unknown>;
    const candidates = [c['preferred_username'], c['email'], c['upn'], c['name']];
    for (const candidate of candidates) {
      if (typeof candidate === 'string' && candidate.length > 0) return candidate;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

function isTokenResponse(value: unknown): value is TokenResponse {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return typeof v['access_token'] === 'string' && typeof v['expires_in'] === 'number';
}

/**
 * Parses a cached `scope` string (the value Azure echoed back at login) into
 * a scope list suitable for a refresh request.
 *
 * Returns `undefined` when the input is absent or empty so the caller can fall
 * through to the `scopesForMode` default.  Otherwise ensures `offline_access`
 * is present — Azure strips it from the echoed scope string even when the
 * original request included it, so a refresh without it would lose the rotated
 * refresh token.
 */
function scopesFromCachedScope(scope: string | undefined): readonly string[] | undefined {
  if (!scope) return undefined;
  const parts = scope.split(/\s+/).filter(Boolean);
  if (parts.length === 0) return undefined;
  if (!parts.includes('offline_access')) parts.unshift('offline_access');
  return parts;
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
