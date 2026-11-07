/**
 * Shared fixtures for the OUTLOOK_POLICY tests: a per-test temp directory for
 * policy and deny-list files, and a fake Graph transport that answers only the
 * sender lookup and the move, recording every request it sees.
 */
import { afterEach, beforeEach, expect, vi } from 'vitest';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { GraphClient } from './client.js';
import { makeCtx } from './commands/testUtils.js';
import { PolicyRefusedError } from './errors.js';

export const REVIEW = 'AAMkReviewFolderId=';
export const ARCHIVE = 'archive';

let dir = '';

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'outlook-cli-policy-'));
});

afterEach(async () => {
  await rm(dir, { recursive: true, force: true });
});

export function tempDir(): string {
  return dir;
}

export async function writeJson(name: string, value: unknown): Promise<string> {
  const path = join(dir, name);
  await writeFile(path, typeof value === 'string' ? value : JSON.stringify(value));
  return path;
}

/** Two destinations; moves into REVIEW are checked against protected.json. */
export const BASE_POLICY = {
  allowCommands: ['move', 'mark-read', 'draft create'],
  move: {
    allowedDestinations: [REVIEW, ARCHIVE],
    denySenders: { file: 'protected.json', maxAgeHours: 24, destinations: [REVIEW] },
  },
};

/** Fake Graph: sender lookup + move. Anything else is a test failure. */
export function fakeGraph(sender: string | undefined) {
  const calls: { method: string; url: string }[] = [];
  const fetchImpl = vi.fn(async (url: string | URL, init?: RequestInit) => {
    const u = String(url);
    const method = init?.method ?? 'GET';
    calls.push({ method, url: u });
    if (method === 'GET' && u.includes('$select=from')) {
      const body = sender ? { from: { emailAddress: { address: sender } } } : {};
      return new Response(JSON.stringify(body), { status: 200 });
    }
    if (method === 'POST' && u.endsWith('/move')) {
      return new Response(JSON.stringify({ id: 'NEW-ID' }), { status: 201 });
    }
    throw new Error(`unexpected request in test: ${method} ${u}`);
  });
  const client = new GraphClient({
    getToken: async () => 't',
    fetchImpl: fetchImpl as unknown as typeof fetch,
  });
  return { client, calls, fetchImpl };
}

/** A command context with OUTLOOK_POLICY pointing at a fresh policy file. */
export async function setup(
  argv: string[],
  command: string,
  opts: { policy?: unknown; sender?: string } = {},
) {
  const policyPath = await writeJson('policy.json', opts.policy ?? BASE_POLICY);
  const t = makeCtx(argv, { OUTLOOK_POLICY: policyPath });
  const graph = fakeGraph(opts.sender);
  t.ctx.client = graph.client;
  t.ctx.command = command;
  return { ...t, ...graph, policyPath };
}

export async function freshDenyList(addresses: string[] = ['boss@example.com']): Promise<string> {
  return writeJson('protected.json', { addresses, built_at: new Date().toISOString() });
}

export async function expectRefused(
  promise: Promise<unknown>,
  rule: string,
): Promise<PolicyRefusedError> {
  const error = await promise.then(
    () => undefined,
    (e: unknown) => e,
  );
  expect(error).toBeInstanceOf(PolicyRefusedError);
  const refused = error as PolicyRefusedError;
  expect(refused.exitCode).toBe(4);
  expect(refused.rule).toBe(rule);
  return refused;
}
