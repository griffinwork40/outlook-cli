/**
 * Tests for `outlook rules create`. Split from rules.test.ts to stay under
 * the 350-LOC file-size gate.
 */

import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import { rulesCommand, type MessageRule } from './rules.js';
import { makeCtx, text } from './testUtils.js';

const CREATED_RULE: MessageRule = {
  id: 'rule-new',
  displayName: 'Test Rule',
  sequence: 3,
  isEnabled: true,
  conditions: {
    fromAddresses: [{ emailAddress: { address: 'spam@example.com' } }],
  },
  actions: {
    moveToFolder: 'AAMkJunk',
    markAsRead: true,
    stopProcessingRules: true,
  },
};

describe('rules create', () => {
  it('POSTs to /me/mailFolders/inbox/messageRules with fromAddresses', async () => {
    const { ctx, client, out } = makeCtx([
      'create', '--name', 'Test Rule', '--folder', 'AAMkJunk',
      '--from', 'spam@example.com', '--mark-read',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    expect(post).toHaveBeenCalledTimes(1);
    expect(post.mock.calls[0]?.[0]).toBe('/me/mailFolders/inbox/messageRules');
    const body = post.mock.calls[0]?.[1] as Record<string, unknown>;
    expect(body.displayName).toBe('Test Rule');
    expect(body.conditions).toEqual({
      fromAddresses: [{ emailAddress: { address: 'spam@example.com' } }],
    });
    expect(body.actions).toEqual({
      moveToFolder: 'AAMkJunk',
      markAsRead: true,
      stopProcessingRules: true,
    });
    expect(body.isEnabled).toBe(true);
    expect(text(out)).toContain('Test Rule');
  });

  it('supports --sender-contains for domain-level matching', async () => {
    const { ctx, client } = makeCtx([
      'create', '--name', 'Domain Rule', '--folder', 'AAMkFolder',
      '--sender-contains', '@example.com',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    const body = post.mock.calls[0]?.[1] as Record<string, unknown>;
    expect(body.conditions).toEqual({ senderContains: ['@example.com'] });
  });

  it('supports multiple --from addresses', async () => {
    const { ctx, client } = makeCtx([
      'create', '--name', 'Multi', '--folder', 'AAMkFolder',
      '--from', 'a@b.com', '--from', 'c@d.com',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    const body = post.mock.calls[0]?.[1] as Record<string, unknown>;
    const conds = body?.conditions as { fromAddresses: Array<{ emailAddress: { address: string } }> };
    expect(conds.fromAddresses).toHaveLength(2);
    expect(conds.fromAddresses[0]?.emailAddress.address).toBe('a@b.com');
    expect(conds.fromAddresses[1]?.emailAddress.address).toBe('c@d.com');
  });

  it('defaults stopProcessingRules to true', async () => {
    const { ctx, client } = makeCtx([
      'create', '--name', 'Stop', '--folder', 'AAMkFolder', '--from', 'x@y.com',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    const body = post.mock.calls[0]?.[1] as Record<string, unknown>;
    const actions = body?.actions as { stopProcessingRules: boolean };
    expect(actions.stopProcessingRules).toBe(true);
  });

  it('respects --no-stop-processing', async () => {
    const { ctx, client } = makeCtx([
      'create', '--name', 'No Stop', '--folder', 'AAMk', '--from', 'x@y.com',
      '--no-stop-processing',
    ]);
    const post = vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    const body = post.mock.calls[0]?.[1] as Record<string, unknown>;
    const actions = body?.actions as { stopProcessingRules?: boolean };
    expect(actions.stopProcessingRules).toBeUndefined();
  });

  it('throws UsageError when --name is missing', async () => {
    const { ctx } = makeCtx(['create', '--folder', 'AAMk', '--from', 'a@b.com']);
    await expect(rulesCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError when --folder is missing', async () => {
    const { ctx } = makeCtx(['create', '--name', 'T', '--from', 'a@b.com']);
    await expect(rulesCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError when no condition is provided', async () => {
    const { ctx } = makeCtx(['create', '--name', 'T', '--folder', 'AAMk']);
    await expect(rulesCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('--dry-run does not touch the network', async () => {
    const { ctx, fetchImpl, out } = makeCtx([
      'create', '--name', 'Dry', '--folder', 'AAMk', '--from', 'a@b.com',
      '--dry-run',
    ]);

    await rulesCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
    expect(text(out)).toContain('DRY RUN');
  });

  it('--dry-run --json emits the request spec', async () => {
    const { ctx, fetchImpl, out } = makeCtx([
      'create', '--name', 'D', '--folder', 'AAMk', '--from', 'a@b.com',
      '--dry-run', '--json',
    ]);

    await rulesCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
    const parsed = JSON.parse(text(out));
    expect(parsed.dryRun).toBe(true);
    expect(parsed.request.method).toBe('POST');
    expect(parsed.request.url).toContain('messageRules');
  });

  it('--json emits the created rule object', async () => {
    const { ctx, client, out } = makeCtx([
      'create', '--name', 'J', '--folder', 'AAMk', '--from', 'a@b.com', '--json',
    ]);
    vi.spyOn(client, 'post').mockResolvedValue(CREATED_RULE);

    await rulesCommand.handler(ctx);

    const parsed = JSON.parse(text(out));
    expect(parsed.id).toBe('rule-new');
    expect(parsed.displayName).toBe('Test Rule');
  });
});
