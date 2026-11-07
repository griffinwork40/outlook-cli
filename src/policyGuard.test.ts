/**
 * OUTLOOK_POLICY enforcement through the real guardMutation chokepoint.
 *
 * Commands run against a GraphClient whose fetch is a fake router (see
 * policyTestUtils.ts) that records every call, so "refused before the request
 * left" is checked against real traffic, not a spy.
 */
import { describe, expect, it, vi } from 'vitest';
import { join } from 'node:path';
import { guardMutation } from './command.js';
import { draftCommand } from './commands/draft.js';
import { getCommand } from './commands/get.js';
import { deleteCommand, moveCommand } from './commands/move.js';
import { sendCommand } from './commands/send.js';
import { makeCtx } from './commands/testUtils.js';
import { ARCHIVE, expectRefused, setup, tempDir } from './policyTestUtils.js';

describe('no policy configured', () => {
  it('guardMutation runs exec exactly as before', async () => {
    const { ctx } = makeCtx([]);
    const exec = vi.fn(async () => 'ok');
    await expect(
      guardMutation(ctx, { method: 'POST', url: 'https://graph.microsoft.com/v1.0/me/sendMail' }, exec),
    ).resolves.toBe('ok');
    expect(exec).toHaveBeenCalledTimes(1);
  });
});

describe('policy file fails closed', () => {
  for (const [label, content] of [
    ['invalid JSON', '{ not json'],
    ['unknown key (typo)', { allowCommands: ['move'], move: { allowedDestination: [ARCHIVE] } }],
    ['missing allowCommands', { move: {} }],
  ] as const) {
    it(`${label}: every mutation is refused and nothing is sent`, async () => {
      const { ctx, calls } = await setup(['MSG1', '--folder', ARCHIVE], 'move', { policy: content });
      await expectRefused(moveCommand.handler(ctx), 'policy-file');
      expect(calls).toHaveLength(0);
    });
  }

  it('unreadable (missing) policy file refuses mutations but not reads', async () => {
    const { ctx, calls } = await setup(['MSG1', '--folder', ARCHIVE], 'move');
    ctx.env = { OUTLOOK_POLICY: join(tempDir(), 'does-not-exist.json') };
    await expectRefused(moveCommand.handler(ctx), 'policy-file');
    expect(calls).toHaveLength(0);

    const read = makeCtx(['MSG1'], { OUTLOOK_POLICY: join(tempDir(), 'does-not-exist.json') });
    read.ctx.command = 'get';
    const get = vi.spyOn(read.client, 'get').mockResolvedValue({ id: 'MSG1', subject: 's' } as never);
    await getCommand.handler(read.ctx);
    expect(get).toHaveBeenCalledTimes(1);
  });
});

describe('allowCommands', () => {
  it('refuses a mutating command that is not listed (send)', async () => {
    const { ctx, calls } = await setup(['--to', 'a@example.com', '--subject', 'S', '--body', 'B'], 'send');
    const error = await expectRefused(sendCommand.handler(ctx), 'allowCommands');
    expect(error.message).toContain("'send' is not in allowCommands");
    expect(calls).toHaveLength(0);
  });

  it('"draft create" allows draft create but not draft send; bare "draft" allows both', async () => {
    const create = await setup(['create', '--subject', 'S', '--no-signature'], 'draft');
    const post = vi.spyOn(create.client, 'post').mockResolvedValue({ id: 'D1' } as never);
    await draftCommand.handler(create.ctx);
    expect(post).toHaveBeenCalledTimes(1);

    const send = await setup(['send', 'D1'], 'draft');
    await expectRefused(draftCommand.handler(send.ctx), 'allowCommands');

    const bare = await setup(['send', 'D1'], 'draft', { policy: { allowCommands: ['draft'] } });
    const sent = vi.spyOn(bare.client, 'post').mockResolvedValue(undefined as never);
    await draftCommand.handler(bare.ctx);
    expect(sent).toHaveBeenCalledTimes(1);
  });

  it('refuses when the running command is unknown to the guard', async () => {
    const { ctx } = await setup(['MSG1', '--folder', ARCHIVE], 'move');
    delete ctx.command;
    await expectRefused(moveCommand.handler(ctx), 'allowCommands');
  });
});

describe('move.allowedDestinations', () => {
  it('allows a listed well-known name (case-insensitive) without any sender lookup', async () => {
    const { ctx, calls } = await setup(['MSG1', '--folder', 'Archive'], 'move');
    await moveCommand.handler(ctx);
    expect(calls.map((c) => c.method)).toEqual(['POST']);
  });

  it('refuses a destination that is not listed', async () => {
    const { ctx, calls } = await setup(['MSG1', '--folder', 'inbox'], 'move');
    await expectRefused(moveCommand.handler(ctx), 'move.allowedDestinations');
    expect(calls).toHaveLength(0);
  });

  it('applies to delete (a move to deleteditems) when delete is allowed as a command', async () => {
    const { ctx } = await setup(['MSG1'], 'delete', {
      policy: { allowCommands: ['delete'], move: { allowedDestinations: [ARCHIVE] } },
    });
    await expectRefused(deleteCommand.handler(ctx), 'move.allowedDestinations');
  });

  it('refuses a message id that would rewrite the request path', async () => {
    const { ctx, calls } = await setup(['../mailFolders/inbox', '--folder', ARCHIVE], 'move');
    await expectRefused(moveCommand.handler(ctx), 'request-path');
    expect(calls).toHaveLength(0);
  });

  it('refuses a double-encoded traversal segment in a move message id', async () => {
    // %252e%252e decodes once to %2e%2e, which still contains '%' — rejected.
    const { ctx, calls } = await setup(['%252e%252e/mailFolders/inbox', '--folder', ARCHIVE], 'move');
    await expectRefused(moveCommand.handler(ctx), 'request-path');
    expect(calls).toHaveLength(0);
  });

  it('purge (DELETE) is refused when deleteditems is not in allowedDestinations', async () => {
    const { ctx, calls } = await setup(['MSG1', '--yes'], 'purge', {
      policy: { allowCommands: ['purge'], move: { allowedDestinations: [ARCHIVE] } },
    });
    const { purgeCommand } = await import('./commands/move.js');
    await expectRefused(purgeCommand.handler(ctx), 'move.allowedDestinations');
    expect(calls).toHaveLength(0);
  });
});
