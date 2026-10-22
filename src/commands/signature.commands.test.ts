/**
 * Signature behaviour at each compose surface: send, draft create/update,
 * reply-family. Unit behaviour lives in src/signature.test.ts.
 */

import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import type { Message } from '../types.js';
import { draftCommand } from './draft.js';
import { forwardCommand, replyCommand } from './reply.js';
import { sendCommand } from './send.js';
import { makeCtx, text } from './testUtils.js';

let env: Record<string, string>;
const SEND = ['--to', 'a@example.com', '--subject', 'S', '--body', 'Hello'];

beforeAll(async () => {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-sigcmd-'));
  const sig = join(dir, 'sig.txt');
  await writeFile(sig, 'Griffin\n');
  env = { OUTLOOK_SIGNATURE_FILE: sig };
});

describe('send with a signature', () => {
  it('appends the configured signature and reports it', async () => {
    const { ctx, client, out } = makeCtx([...SEND, '--json'], env);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { body: { content: string } };
    };
    expect(payload.message.body.content).toBe('Hello\n\n-- \nGriffin');
    expect(JSON.parse(text(out))).toMatchObject({ signatureApplied: true });
  });

  it('--no-signature sends the body untouched', async () => {
    const { ctx, client, out } = makeCtx([...SEND, '--no-signature', '--json'], env);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { body: { content: string } };
    };
    expect(payload.message.body.content).toBe('Hello');
    expect(JSON.parse(text(out))).toMatchObject({ signatureApplied: false });
  });

  it('shows the signature in --dry-run without touching the network', async () => {
    const { ctx, out, fetchImpl } = makeCtx([...SEND, '--dry-run'], env);

    await sendCommand.handler(ctx);

    expect(fetchImpl).not.toHaveBeenCalled();
    expect(text(out)).toContain('-- \\nGriffin');
  });

  it('nothing configured means no signature (hermetic default)', async () => {
    const { ctx, client } = makeCtx(SEND);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { body: { content: string } };
    };
    expect(payload.message.body.content).toBe('Hello');
  });
});

describe('draft with a signature', () => {
  const DRAFT: Message = { id: 'D1', subject: 'S', isDraft: true };

  it('create signs the draft', async () => {
    const { ctx, client } = makeCtx(['create', '--subject', 'S', '--body', 'B'], env);
    const post = vi.spyOn(client, 'post').mockResolvedValue(DRAFT as never);

    await draftCommand.handler(ctx);

    const message = post.mock.calls[0]?.[1] as { body: { content: string } };
    expect(message.body.content).toBe('B\n\n-- \nGriffin');
  });

  it('create with nothing to say still refuses, even with a signature configured', async () => {
    const { ctx } = makeCtx(['create'], env);
    await expect(draftCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });

  it('update never appends the signature', async () => {
    const { ctx, client } = makeCtx(['update', 'D1', '--body', 'New'], env);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue(DRAFT as never);

    await draftCommand.handler(ctx);

    const message = patch.mock.calls[0]?.[1] as { body: { content: string } };
    expect(message.body.content).toBe('New');
  });

  it('update rejects an explicit --signature request', async () => {
    const { ctx } = makeCtx(['update', 'D1', '--body', 'New', '--signature'], env);
    await expect(draftCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
  });
});

describe('reply-family with a signature', () => {
  it('item 2: reply with no --body sends only the signature block', async () => {
    const { ctx, client } = makeCtx(['MSG1'], env);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyCommand.handler(ctx);

    expect(post.mock.calls[0]?.[1]).toEqual({ comment: '-- \nGriffin' });
  });

  it('reply appends the text signature to the comment', async () => {
    const { ctx, client } = makeCtx(['MSG1', '--body', 'Thanks!'], env);
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await replyCommand.handler(ctx);

    expect(post.mock.calls[0]?.[1]).toEqual({ comment: 'Thanks!\n\n-- \nGriffin' });
  });

  it('item 5: send with missing sig file and no --to reports recipient error', async () => {
    const { ctx } = makeCtx(
      ['--subject', 'S', '--body', 'Hi'],
      { OUTLOOK_SIGNATURE_FILE: '/nonexistent/missing-sig.txt' },
    );
    const err = await sendCommand.handler(ctx).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain('recipient');
  });

  it('item 5: draft create with no content and missing sig file reports nothing-to-create', async () => {
    const { ctx } = makeCtx(
      ['create'],
      { OUTLOOK_SIGNATURE_FILE: '/nonexistent/missing-sig.txt' },
    );
    const err = await draftCommand.handler(ctx).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain('Nothing to create');
  });

  it('forward honours --no-signature', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1', '--to', 'x@example.com', '--body', 'FYI', '--no-signature'],
      env,
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await forwardCommand.handler(ctx);

    expect(post.mock.calls[0]?.[1]).toMatchObject({ comment: 'FYI' });
  });
});

describe('HTML-only signature config — plain-text surfaces reject loudly (F5 / R2.4)', () => {
  let htmlOnlyEnv: Record<string, string>;

  beforeAll(async () => {
    const dir = await mkdtemp(join(tmpdir(), 'outlook-htmlsig-'));
    const htmlSig = join(dir, 'sig.html');
    await writeFile(htmlSig, '<b>Griffin</b>');
    htmlOnlyEnv = { OUTLOOK_SIGNATURE_HTML_FILE: htmlSig };
  });

  it('F5: send without --html throws UsageError mentioning OUTLOOK_SIGNATURE_FILE', async () => {
    const { ctx } = makeCtx(
      ['--to', 'a@example.com', '--subject', 'S', '--body', 'Hello'],
      htmlOnlyEnv,
    );
    const err = await sendCommand.handler(ctx).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain('OUTLOOK_SIGNATURE_FILE');
  });

  it('R2.4: reply with no --body throws UsageError mentioning OUTLOOK_SIGNATURE_FILE', async () => {
    const { ctx } = makeCtx(['MSG1'], htmlOnlyEnv);
    const err = await replyCommand.handler(ctx).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UsageError);
    expect((err as Error).message).toContain('OUTLOOK_SIGNATURE_FILE');
  });
});
