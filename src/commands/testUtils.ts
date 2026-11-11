/**
 * Shared test scaffolding for command tests.
 *
 * Deliberately builds a REAL GraphClient (so `buildSpec` behaves exactly as in
 * production) but injects a fetch that THROWS. Any command that reaches the
 * network when it should not — notably under --dry-run — fails loudly instead
 * of quietly passing.
 */

import { vi } from 'vitest';
import { parseArgs } from '../args.js';
import { GraphClient } from '../client.js';
import type { Authenticator } from '../auth.js';
import type { CommandContext } from '../command.js';

export function makeClient(): {
  client: GraphClient;
  fetchImpl: ReturnType<typeof vi.fn>;
} {
  const fetchImpl = vi.fn(async () => {
    throw new Error('NETWORK CALLED — a guarded path leaked to the network');
  });
  const client = new GraphClient({
    getToken: async () => 'test-token',
    fetchImpl: fetchImpl as unknown as typeof fetch,
  });
  return { client, fetchImpl };
}

const stubAuth: Authenticator = {
  getAccessToken: async () => 'test-token',
  login: async () => {
    throw new Error('not used');
  },
  logout: async () => undefined,
  refresh: async () => {
    throw new Error('not used');
  },
  status: async () => ({ signedIn: true, expired: false }),
};

export interface TestCtx {
  ctx: CommandContext;
  out: string[];
  err: string[];
  client: GraphClient;
  fetchImpl: ReturnType<typeof vi.fn>;
}

export function makeCtx(
  argv: string[],
  env: Record<string, string | undefined> = {},
  options: { command?: string } = {},
): TestCtx {
  const out: string[] = [];
  const err: string[] = [];
  const { client, fetchImpl } = makeClient();
  const args = parseArgs(argv);

  const ctx: CommandContext = {
    args,
    json: args.bool('json'),
    dryRun: args.bool('dry-run'),
    ...(options.command !== undefined ? { command: options.command } : {}),
    client,
    auth: stubAuth,
    env,
    out: (line) => out.push(line),
    err: (line) => err.push(line),
  };

  return { ctx, out, err, client, fetchImpl };
}

/** Joined stdout, for substring assertions. */
export function text(out: string[]): string {
  return out.join('\n');
}
