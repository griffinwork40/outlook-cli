/**
 * refresh() uses cached.scope to preserve scopes that were granted at login
 * but are absent from the current binary's DEFAULT_SCOPES / READONLY_SCOPES.
 *
 * Covers issue #16: a token obtained under a fork binary (e.g. feat/contacts
 * which included Contacts.ReadWrite) must keep those scopes across refresh
 * even after switching to a main build whose DEFAULT_SCOPES differ.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createAuthenticator } from './auth.js';
import { createTokenStore } from './tokenStore.js';
import type { OutlookConfig } from './types.js';

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

function refreshResponse(extra?: object) {
  return jsonResponse({
    token_type: 'Bearer',
    expires_in: 3600,
    access_token: 'new-access',
    refresh_token: 'new-refresh',
    ...extra,
  });
}

/** Captures the `scope` form-field of every /token call. */
function capturingFetch() {
  const requestedScopes: string[][] = [];
  const fetchImpl = vi.fn().mockImplementation(async (_url: string | URL, init?: RequestInit) => {
    const scope = new URLSearchParams(init?.body as string).get('scope') ?? '';
    requestedScopes.push(scope.split(/\s+/).filter(Boolean));
    return refreshResponse();
  });
  return { fetchImpl, requestedScopes };
}

describe('refresh() — cached.scope fallback (issue #16)', () => {
  let dir: string;
  let tokenPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'outlook-cli-cached-scope-'));
    tokenPath = join(dir, 'token.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('a token with extra scopes in cached.scope keeps them across refresh', async () => {
    // Simulates a token obtained under a fork binary (e.g. feat/contacts) that
    // added Contacts.ReadWrite. No grantScopes field (predates --scopes login).
    // The current binary's DEFAULT_SCOPES do not include Contacts.ReadWrite.
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT',
      expiresAt: 1_000,
      mode: 'full',
      scope: 'User.Read Mail.ReadWrite Mail.Send Contacts.ReadWrite Contacts.Read',
      // no grantScopes — older token format
    });

    const { fetchImpl, requestedScopes } = capturingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, now: () => 2_000 });

    await auth.refresh();

    // offline_access must be prepended (Azure strips it on echo)
    expect(requestedScopes[0]).toContain('offline_access');
    // The extra Contacts scopes must be preserved — not silently dropped
    expect(requestedScopes[0]).toContain('Contacts.ReadWrite');
    expect(requestedScopes[0]).toContain('Contacts.Read');
    expect(requestedScopes[0]).toContain('Mail.ReadWrite');
  });

  it('a readonly token stays readonly across refresh — cached.scope cannot widen the grant', async () => {
    // The readonly grant's echoed scope will only contain the read-only scopes,
    // so using cached.scope here cannot accidentally re-escalate to full access.
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_RO',
      expiresAt: 1_000,
      mode: 'readonly',
      scope: 'User.Read Mail.Read MailboxSettings.Read',
    });

    const { fetchImpl, requestedScopes } = capturingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, now: () => 2_000 });

    await auth.refresh();
    await auth.refresh(); // second refresh — must still be readonly

    for (const scopes of requestedScopes) {
      expect(scopes).toContain('offline_access');
      expect(scopes).toContain('Mail.Read');
      expect(scopes).not.toContain('Mail.ReadWrite');
      expect(scopes).not.toContain('Mail.Send');
    }
  });

  it('grantScopes wins over cached.scope when both are present', async () => {
    // A --scopes login sets grantScopes; that must win over the echoed scope.
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT',
      expiresAt: 1_000,
      mode: 'full',
      scope: 'User.Read Mail.ReadWrite Mail.Send Contacts.ReadWrite',
      grantScopes: ['offline_access', 'Mail.ReadWrite'],
    });

    const { fetchImpl, requestedScopes } = capturingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, now: () => 2_000 });

    await auth.refresh();

    expect(requestedScopes[0]).toEqual(['offline_access', 'Mail.ReadWrite']);
    expect(requestedScopes[0]).not.toContain('Contacts.ReadWrite');
  });

  it('falls back to scopesForMode when cached.scope is absent', async () => {
    // Legacy token with no scope field and no grantScopes.
    const store = createTokenStore(tokenPath);
    await store.save({
      accessToken: 'old-access',
      refreshToken: 'RT_LEGACY',
      expiresAt: 1_000,
      mode: 'full',
      // no scope, no grantScopes
    });

    const { fetchImpl, requestedScopes } = capturingFetch();
    const auth = createAuthenticator(cfg, store, { fetchImpl, now: () => 2_000 });

    await auth.refresh();

    // Should use DEFAULT_SCOPES (scopesForMode('full'))
    expect(requestedScopes[0]).toContain('offline_access');
    expect(requestedScopes[0]).toContain('Mail.ReadWrite');
    expect(requestedScopes[0]).toContain('Mail.Send');
  });
});
