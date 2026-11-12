/**
 * OUTLOOK_POLICY: the sender deny list on protected destinations, --dry-run
 * reporting, and the --json refusal shape.
 */
import { describe, expect, it } from 'vitest';
import { utimes } from 'node:fs/promises';
import { moveCommand } from './commands/move.js';
import { text } from './commands/testUtils.js';
import { validatePolicy } from './policy.js';
import {
  ARCHIVE,
  REVIEW,
  expectRefused,
  freshDenyList,
  setup,
  writeJson,
} from './policyTestUtils.js';

const HOUR = 3_600_000;

describe('move.denySenders', () => {
  it('refuses moving a message whose sender is on the list, after a $select=from lookup', async () => {
    await freshDenyList(['boss@example.com']);
    const { ctx, calls } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'Boss@Example.com' });

    const error = await expectRefused(moveCommand.handler(ctx), 'move.denySenders');

    expect(error.message).toContain('boss@example.com');
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe('GET');
    expect(calls[0]?.url).toContain('/me/messages/MSG1?$select=from');
  });

  it('allows the move when the sender is not on the list', async () => {
    await freshDenyList(['boss@example.com']);
    const { ctx, calls } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'news@example.org' });

    await moveCommand.handler(ctx);

    expect(calls.map((c) => c.method)).toEqual(['GET', 'POST']);
  });

  it('accepts a plain array deny list (age taken from mtime)', async () => {
    await writeJson('protected.json', ['boss@example.com']);
    const { ctx } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'boss@example.com' });
    await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
  });

  it('does not consult the list for a destination it does not protect', async () => {
    const { ctx, calls } = await setup(['MSG1', '--folder', ARCHIVE], 'move', { sender: 'boss@example.com' });
    await moveCommand.handler(ctx); // no protected.json exists at all
    expect(calls.map((c) => c.method)).toEqual(['POST']);
  });

  it('missing deny list: moves to protected destinations are refused without any request', async () => {
    const { ctx, calls } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    const error = await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
    expect(error.message).toContain('cannot read deny list');
    expect(calls).toHaveLength(0);
  });

  it('unparseable deny list is refused', async () => {
    await writeJson('protected.json', '{ nope');
    const { ctx } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
  });

  it('stale by built_at is refused', async () => {
    await writeJson('protected.json', {
      addresses: ['stale@example.com'],
      built_at: new Date(Date.now() - 25 * HOUR).toISOString(),
    });
    const { ctx, calls } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    const error = await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
    expect(error.message).toContain('old');
    expect(calls).toHaveLength(0);
  });

  it('stale by mtime (no built_at) is refused', async () => {
    const file = await writeJson('protected.json', ['stale@example.com']);
    const old = new Date(Date.now() - 48 * HOUR);
    await utimes(file, old, old);
    const { ctx } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
  });

  it('a message with no sender address is refused for a protected destination', async () => {
    await freshDenyList();
    const { ctx, calls } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: undefined });
    await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
    expect(calls.map((c) => c.method)).toEqual(['GET']);
  });
});

describe('--dry-run with a policy', () => {
  it('prints the request, reports a refusal, and makes no network call', async () => {
    const { ctx, calls, out } = await setup(['MSG1', '--folder', 'inbox', '--dry-run'], 'move');

    await moveCommand.handler(ctx);

    expect(calls).toHaveLength(0);
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('would REFUSE (move.allowedDestinations)');
  });

  it('JSON: reports allowed with the sender check skipped for a protected destination', async () => {
    await freshDenyList();
    const { ctx, calls, out } = await setup(['MSG1', '--folder', REVIEW, '--dry-run', '--json'], 'move', {
      sender: 'boss@example.com',
    });

    await moveCommand.handler(ctx);

    expect(calls).toHaveLength(0);
    const parsed = JSON.parse(out[0] ?? '') as {
      dryRun: boolean;
      policy: { allowed: boolean; senderCheck?: string };
    };
    expect(parsed.dryRun).toBe(true);
    expect(parsed.policy.allowed).toBe(true);
    expect(parsed.policy.senderCheck).toBe('skipped-dry-run');
  });

  it('JSON: a stale deny list shows up as a refusal in the preview', async () => {
    await writeJson('protected.json', { addresses: ['stale@example.com'], built_at: '2000-01-01T00:00:00Z' });
    const { ctx, out } = await setup(['MSG1', '--folder', REVIEW, '--dry-run', '--json'], 'move');

    await moveCommand.handler(ctx);

    const parsed = JSON.parse(out[0] ?? '') as { policy: { allowed: boolean; rule: string } };
    expect(parsed.policy).toMatchObject({ allowed: false, rule: 'move.denySenders' });
  });
});

describe('--json refusal', () => {
  it('writes a JSON error object to stdout and throws with exit code 4', async () => {
    const { ctx, out } = await setup(['MSG1', '--folder', 'inbox', '--json'], 'move');

    await expectRefused(moveCommand.handler(ctx), 'move.allowedDestinations');

    const parsed = JSON.parse(out[0] ?? '') as { error: { code: string; rule: string; message: string } };
    expect(parsed.error.code).toBe('PolicyRefused');
    expect(parsed.error.rule).toBe('move.allowedDestinations');
    expect(parsed.error.message).toContain('inbox');
  });
});

describe('validatePolicy — deny.file path traversal', () => {
  it('rejects an absolute path', () => {
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: '/etc/passwd', maxAgeHours: 1, destinations: ['a'] } },
      }),
    ).toContain('relative path');
  });

  it('rejects a path with .. segments', () => {
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: '../outside.json', maxAgeHours: 1, destinations: ['a'] } },
      }),
    ).toContain('relative path');
  });

  it('rejects a path with backslashes', () => {
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: 'sub\\file.json', maxAgeHours: 1, destinations: ['a'] } },
      }),
    ).toContain('relative path');
  });

  it('allows a nested relative path without .. or backslashes', () => {
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: 'sub/deny.json', maxAgeHours: 1, destinations: ['a'] } },
      }),
    ).toBeUndefined();
  });
});

describe('validatePolicy', () => {
  it('accepts the documented shape, with $comment keys', () => {
    expect(
      validatePolicy({
        $comment: 'agent policy',
        allowCommands: ['move'],
        move: {
          allowedDestinations: ['a'],
          denySenders: { file: 'p.json', maxAgeHours: 1, destinations: ['a'] },
        },
      }),
    ).toBeUndefined();
  });

  it('rejects bad field types', () => {
    expect(validatePolicy([])).toBeDefined();
    expect(validatePolicy({ allowCommands: 'move' })).toBeDefined();
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: 'p.json', maxAgeHours: 0, destinations: ['a'] } },
      }),
    ).toContain('maxAgeHours');
    expect(
      validatePolicy({
        allowCommands: [],
        move: { denySenders: { file: 'p.json', maxAgeHours: 1, destinations: [] } },
      }),
    ).toContain('destinations');
  });
});

describe('loadDenyList — empty addresses (finding 6: fail-closed)', () => {
  it('rejects a plain empty array (fail-closed: empty list cannot vouch a sender is absent)', async () => {
    await writeJson('protected.json', []);
    const { ctx } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    const error = await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
    expect(error.message).toContain('non-empty');
  });

  it('rejects an object with an empty addresses array', async () => {
    await writeJson('protected.json', { addresses: [], built_at: new Date().toISOString() });
    const { ctx } = await setup(['MSG1', '--folder', REVIEW], 'move', { sender: 'x@example.org' });
    const error = await expectRefused(moveCommand.handler(ctx), 'move.denySenders');
    expect(error.message).toContain('non-empty');
  });
});
