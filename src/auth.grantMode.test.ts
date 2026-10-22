/**
 * Read-only grant mode (`outlook auth login --readonly`).
 *
 * Split out of auth.test.ts purely to stay under the 350-LOC file ceiling —
 * these exercise the same createAuthenticator surface, scoped to how a grant's
 * scope set is chosen at login and then PRESERVED across refreshes.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createAuthenticator } from './auth.js';
import { createTokenStore } from './tokenStore.js';
import { NotAuthenticatedError } from './errors.js';
import type { OutlookConfig, TokenSet } from './types.js';

const cfg: OutlookConfig = {
  clientId: 'client-abc-123',
  tenantId: 'common',
  tokenCachePath: '/dev/null',
};

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/** Encodes a fake (unsigned, header/payload/sig) id_token carrying the given claims. */
function fakeIdToken(claims: Record<string, unknown>): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `${header}.${payload}.`;
}

const TOKEN_BODY = {
  token_type: 'Bearer',
  scope: 'User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite',
  expires_in: 3599,
  access_token: 'ACCESS_TOKEN_VALUE',
  refresh_token: 'REFRESH_TOKEN_VALUE',
  id_token: fakeIdToken({ preferred_username: 'user@example.com' }),
};

const DEVICE_CODE_BODY = {
  device_code: 'DEVICECODE_OPAQUE',
  user_code: 'ABCD1234',
  verification_uri: 'https://microsoft.com/devicelogin',
  expires_in: 900,
  interval: 5,
  message: 'To sign in...',
};

function instantSleep(): (ms: number) => Promise<void> {
  return vi.fn().mockResolvedValue(undefined);
}

/** Routes /devicecode vs /token so login() can be driven end-to-end against one mock. */
function loginFlowFetch(tokenBody: unknown = TOKEN_BODY, tokenStatus = 200) {
  return vi.fn().mockImplementation(async (url: string | URL) => {
    const u = String(url);
    if (u.includes('/devicecode')) return jsonResponse(DEVICE_CODE_BODY);
    if (u.includes('/token')) return jsonResponse(tokenBody, tokenStatus);
    throw new Error(`unexpected url in test: ${u}`);
  });
}

describe('createAuthenticator — read-only grant mode', () => {
  // Restated as literals rather than imported from types.ts on purpose: these
  // are the scope sets we intend to request, so an accidental edit to
  // DEFAULT_SCOPES/READONLY_SCOPES must FAIL here rather than be followed.
  const FULL_SCOPE_SET = ['offline_access', 'User.Read', 'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.ReadWrite'];
  const READONLY_SCOPE_SET = ['offline_access', 'User.Read', 'Mail.Read', 'MailboxSettings.Read'];

  let dir: string;
  let tokenPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-auth-test-'));
    tokenPath = join(dir, 'token.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('login(): with no options requests the full scope set and saves mode "full"', async () => {
    const store = createTokenStore(tokenPath);
    let requestedScope = '';
    const fetchMock = vi.fn().mockImplementation(async (url: string | URL, init?: RequestInit) => {
      const u = String(url);
      if (u.includes('/devicecode')) {
        requestedScope = new URLSearchParams(init?.body as string).get('scope') ?? '';
        return jsonResponse(DEVICE_CODE_BODY);
      }
      if (u.includes('/token')) return jsonResponse(TOKEN_BODY);
      throw new Error(`unexpected url in test: ${u}`);
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 1_000_000,
    });

    const tokens = await authenticator.login(() => {});

    // Split on whitespace before asserting scope membership — see the
    // readOnly test below for why a raw substring check is unsafe here too.
    const requestedScopes = requestedScope.split(/\s+/).filter(Boolean);
    expect(requestedScopes).toEqual(FULL_SCOPE_SET);
    expect(requestedScopes).toContain('Mail.ReadWrite');
    expect(requestedScopes).toContain('Mail.Send');
    expect(tokens.mode).toBe('full');

    const persisted = await store.load();
    expect(persisted?.mode).toBe('full');
  });

  it('login(): with { readOnly: true } requests the reduced scope set and saves mode "readonly"', async () => {
    const store = createTokenStore(tokenPath);
    let requestedScope = '';
    const fetchMock = vi.fn().mockImplementation(async (url: string | URL, init?: RequestInit) => {
      const u = String(url);
      if (u.includes('/devicecode')) {
        requestedScope = new URLSearchParams(init?.body as string).get('scope') ?? '';
        return jsonResponse(DEVICE_CODE_BODY);
      }
      if (u.includes('/token')) return jsonResponse(TOKEN_BODY);
      throw new Error(`unexpected url in test: ${u}`);
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 1_000_000,
    });

    const tokens = await authenticator.login(() => {}, { readOnly: true });

    // 'Mail.ReadWrite' contains 'Mail.Read' as a substring — split on
    // whitespace into discrete tokens first, so this can't false-pass the
    // way a naive `.includes()` check on the raw scope string would.
    const requestedScopes = requestedScope.split(/\s+/).filter(Boolean);
    expect(requestedScopes).toEqual(READONLY_SCOPE_SET);
    expect(requestedScopes).toContain('Mail.Read');
    expect(requestedScopes).not.toContain('Mail.ReadWrite');
    expect(requestedScopes).not.toContain('Mail.Send');
    expect(tokens.mode).toBe('readonly');

    const persisted = await store.load();
    expect(persisted?.mode).toBe('readonly');
  });

  it('refresh(): on a cached token with mode "readonly" POSTs the read-only scope set', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_READONLY',
      expiresAt: 1_000_000,
      mode: 'readonly',
    });

    let requestedScope = '';
    const fetchMock = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
      requestedScope = new URLSearchParams(init?.body as string).get('scope') ?? '';
      return jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'new-access',
        refresh_token: 'new-refresh',
      });
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    await authenticator.refresh();

    const requestedScopes = requestedScope.split(/\s+/).filter(Boolean);
    expect(requestedScopes).toEqual(READONLY_SCOPE_SET);
    expect(requestedScopes).not.toContain('Mail.ReadWrite');
  });

  it('refresh(): on a cached token with mode "full" POSTs the full scope set', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_FULL',
      expiresAt: 1_000_000,
      mode: 'full',
    });

    let requestedScope = '';
    const fetchMock = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
      requestedScope = new URLSearchParams(init?.body as string).get('scope') ?? '';
      return jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'new-access',
        refresh_token: 'new-refresh',
      });
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    await authenticator.refresh();

    const requestedScopes = requestedScope.split(/\s+/).filter(Boolean);
    expect(requestedScopes).toEqual(FULL_SCOPE_SET);
    expect(requestedScopes).toContain('Mail.ReadWrite');
    expect(requestedScopes).toContain('Mail.Send');
  });

  it('refresh(): a legacy cached token with no mode field is treated as full and still succeeds', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_LEGACY',
      expiresAt: 1_000_000,
      // no `mode` — predates read-only grants
    });

    let requestedScope = '';
    const fetchMock = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
      requestedScope = new URLSearchParams(init?.body as string).get('scope') ?? '';
      return jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: 'new-access',
        refresh_token: 'new-refresh',
      });
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    const result = await authenticator.refresh();
    expect(result.accessToken).toBe('new-access');

    const requestedScopes = requestedScope.split(/\s+/).filter(Boolean);
    expect(requestedScopes).toEqual(FULL_SCOPE_SET);
  });

  it('refresh(): a "readonly" mode survives refresh, including a second (compounding) refresh', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_1',
      expiresAt: 1_000_000,
      mode: 'readonly',
    });

    const requestedScopesPerCall: string[][] = [];
    let refreshCounter = 0;
    const fetchMock = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
      refreshCounter += 1;
      const scope = new URLSearchParams(init?.body as string).get('scope') ?? '';
      requestedScopesPerCall.push(scope.split(/\s+/).filter(Boolean));
      return jsonResponse({
        token_type: 'Bearer',
        expires_in: 3600,
        access_token: `new-access-${refreshCounter}`,
        refresh_token: `new-refresh-${refreshCounter}`,
      });
    });
    const authenticator = createAuthenticator(cfg, store, {
      fetchImpl: fetchMock,
      sleep: instantSleep(),
      now: () => 2_000_000,
    });

    // First refresh: mode must survive into the persisted TokenSet.
    const firstResult = await authenticator.refresh();
    expect(firstResult.mode).toBe('readonly');
    const persistedAfterFirst = await store.load();
    expect(persistedAfterFirst?.mode).toBe('readonly');

    // Second (compounding) refresh, fed from the token the first refresh just
    // saved: a bug that drops `mode` on refresh would pass the "readonly"
    // POST assertions above but silently escalate here, requesting the full
    // read/write scope set on this second call.
    const secondResult = await authenticator.refresh();
    expect(secondResult.mode).toBe('readonly');
    const persistedAfterSecond = await store.load();
    expect(persistedAfterSecond?.mode).toBe('readonly');

    expect(requestedScopesPerCall).toHaveLength(2);
    expect(requestedScopesPerCall[0] ?? []).toEqual(READONLY_SCOPE_SET);
    expect(requestedScopesPerCall[1] ?? []).toEqual(READONLY_SCOPE_SET);
    expect(requestedScopesPerCall[1] ?? []).not.toContain('Mail.ReadWrite');
  });

  it('status(): surfaces `mode` when the cached token has one, and omits the key entirely when it does not', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'a',
      refreshToken: 'r',
      expiresAt: 2_000_000,
      mode: 'readonly',
    });
    const authenticatorWithMode = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      now: () => 1_000_000,
    });
    const statusWithMode = await authenticatorWithMode.status();
    expect(statusWithMode.mode).toBe('readonly');
    expect(statusWithMode).toHaveProperty('mode');

    await store.save({
      accessToken: 'a',
      refreshToken: 'r',
      expiresAt: 2_000_000,
      // no `mode` — legacy cache predates read-only grants
    });
    const authenticatorWithoutMode = createAuthenticator(cfg, store, {
      fetchImpl: vi.fn(),
      now: () => 1_000_000,
    });
    const statusWithoutMode = await authenticatorWithoutMode.status();
    expect(statusWithoutMode).not.toHaveProperty('mode');
  });
});
