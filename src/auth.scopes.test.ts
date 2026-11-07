/**
 * Scoped login (`outlook auth login --scopes a,b`): the exact scope set is
 * requested, persisted as `grantScopes`, and re-requested on every refresh,
 * so a deliberately narrow token never widens back to DEFAULT_SCOPES.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createAuthenticator } from './auth.js';
import { authCommand } from './commands/auth.js';
import { makeCtx } from './commands/testUtils.js';
import { UsageError } from './errors.js';
import { assertKnownFlags } from './flagSpec.js';
import { parseScopeList } from './scopes.js';
import { createTokenStore } from './tokenStore.js';
import type { OutlookConfig, TokenSet } from './types.js';

const cfg: OutlookConfig = { clientId: 'client-abc', tenantId: 'common', tokenCachePath: '/dev/null' };

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
}

const TOKEN_BODY = {
  token_type: 'Bearer',
  scope: 'Mail.ReadWrite User.Read',
  expires_in: 3599,
  access_token: 'ACCESS',
  refresh_token: 'REFRESH',
};

/** Records the `scope` form field of every identity-platform call. */
function recordingFetch() {
  const scopes: { url: string; scope: string[] }[] = [];
  const fetchImpl = vi.fn().mockImplementation(async (url: string | URL, init?: RequestInit) => {
    const u = String(url);
    const scope = (new URLSearchParams(init?.body as string).get('scope') ?? '').split(/\s+/).filter(Boolean);
    scopes.push({ url: u, scope });
    if (u.includes('/devicecode')) {
      return jsonResponse({
        device_code: 'DC',
        user_code: 'UC',
        verification_uri: 'https://microsoft.com/devicelogin',
        expires_in: 900,
        interval: 5,
      });
    }
    return jsonResponse(TOKEN_BODY);
  });
  return { fetchImpl, scopes };
}

describe('parseScopeList', () => {
  it('splits a comma list, trims, and prepends offline_access when missing', () => {
    expect(parseScopeList(' Mail.ReadWrite , User.Read ')).toEqual([
      'offline_access',
      'Mail.ReadWrite',
      'User.Read',
    ]);
  });

  it('keeps an explicit offline_access in place and drops case-insensitive duplicates', () => {
    expect(parseScopeList('Mail.Read,offline_access,mail.read')).toEqual(['Mail.Read', 'offline_access']);
  });

  it('rejects an empty list and an invalid token', () => {
    expect(() => parseScopeList(' , ')).toThrow(UsageError);
    expect(() => parseScopeList('Mail.Read&x=1')).toThrow(UsageError);
  });
});

describe('createAuthenticator — custom scope set', () => {
  let dir: string;
  let tokenPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-scopes-'));
    tokenPath = join(dir, 'token.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('login({ scopes }) requests exactly that set and persists it as grantScopes', async () => {
    const store = createTokenStore(tokenPath);
    const { fetchImpl, scopes } = recordingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, sleep: async () => undefined, now: () => 1_000 });

    const tokens = await auth.login(() => {}, { scopes: ['offline_access', 'Mail.ReadWrite'] });

    const device = scopes.find((s) => s.url.includes('/devicecode'));
    expect(device?.scope).toEqual(['offline_access', 'Mail.ReadWrite']);
    expect(tokens.grantScopes).toEqual(['offline_access', 'Mail.ReadWrite']);
    const onDisk = JSON.parse(await readFile(tokenPath, 'utf8')) as TokenSet;
    expect(onDisk.grantScopes).toEqual(['offline_access', 'Mail.ReadWrite']);
    expect((await auth.status()).grantScopes).toEqual(['offline_access', 'Mail.ReadWrite']);
  });

  it('refresh() re-requests the cached grantScopes, not DEFAULT_SCOPES, and keeps them', async () => {
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old',
      refreshToken: 'RT',
      expiresAt: 1_000,
      mode: 'full',
      grantScopes: ['offline_access', 'Mail.ReadWrite'],
    });
    const { fetchImpl, scopes } = recordingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, now: () => 2_000 });

    await auth.refresh();
    await auth.refresh();

    expect(scopes).toHaveLength(2);
    for (const call of scopes) {
      expect(call.scope).toEqual(['offline_access', 'Mail.ReadWrite']);
      expect(call.scope).not.toContain('Mail.Send');
    }
    expect((await store.load())?.grantScopes).toEqual(['offline_access', 'Mail.ReadWrite']);
  });

  it('login() without scopes stores no grantScopes (default behavior unchanged)', async () => {
    const store = createTokenStore(tokenPath);
    const { fetchImpl, scopes } = recordingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, sleep: async () => undefined, now: () => 1_000 });

    const tokens = await auth.login(() => {});

    expect(tokens.grantScopes).toBeUndefined();
    expect(scopes[0]?.scope).toContain('Mail.Send');
  });

  it('a cache with a malformed grantScopes is treated as no session (never widened)', async () => {
    await writeFile(
      tokenPath,
      JSON.stringify({ accessToken: 'a', expiresAt: 1, refreshToken: 'r', grantScopes: 'Mail.ReadWrite' }),
    );
    expect(await createTokenStore(tokenPath).load()).toBeNull();
  });
});

describe('outlook auth --scopes (command + flag allowlist)', () => {
  it('the flag allowlist accepts --scopes on auth', () => {
    expect(() => assertKnownFlags('auth', ['scopes'])).not.toThrow();
  });

  it('login --scopes passes the normalized list to the authenticator', async () => {
    const { ctx } = makeCtx(['login', '--scopes', 'Mail.ReadWrite,User.Read', '--json']);
    const login = vi.spyOn(ctx.auth, 'login').mockResolvedValue({
      accessToken: 'secret',
      expiresAt: 5,
      grantScopes: ['offline_access', 'Mail.ReadWrite', 'User.Read'],
    });

    await authCommand.handler(ctx);

    expect(login.mock.calls[0]?.[1]).toEqual({
      readOnly: false,
      scopes: ['offline_access', 'Mail.ReadWrite', 'User.Read'],
    });
  });

  it('login without --scopes passes no scopes option', async () => {
    const { ctx } = makeCtx(['login']);
    const login = vi.spyOn(ctx.auth, 'login').mockResolvedValue({ accessToken: 's', expiresAt: 5 });

    await authCommand.handler(ctx);

    expect(login.mock.calls[0]?.[1]).toEqual({ readOnly: false });
  });

  it('--scopes with --readonly, on refresh, or without a value is a usage error', async () => {
    for (const argv of [
      ['login', '--readonly', '--scopes', 'Mail.Read'],
      ['refresh', '--scopes', 'Mail.Read'],
      ['login', '--scopes'],
    ]) {
      const { ctx } = makeCtx(argv);
      await expect(authCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    }
  });
});
