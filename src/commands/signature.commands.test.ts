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
import { forwardCommand, replyAllCommand, replyCommand } from './reply.js';
import { sendCommand } from './send.js';
import { makeCtx, text } from './testUtils.js';

let env: Record<string, string>;
let htmlEnv: Record<string, string>;
let htmlTextEnv: Record<string, string>;
const SEND = ['--to', 'a@example.com', '--subject', 'S', '--body', 'Hello'];

beforeAll(async () => {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-sigcmd-'));
  const sig = join(dir, 'sig.txt');
  const sigHtml = join(dir, 'sig.html');
  await writeFile(sig, 'Griffin\n');
  await writeFile(sigHtml, '<b>Griffin</b>\n');
  env = { OUTLOOK_SIGNATURE_FILE: sig };
  htmlEnv = { OUTLOOK_SIGNATURE_HTML_FILE: sigHtml };
  htmlTextEnv = { OUTLOOK_SIGNATURE_FILE: sig, OUTLOOK_SIGNATURE_HTML_FILE: sigHtml };
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

describe('reply-family with an HTML signature — three-step flow', () => {
  const HTML_DRAFT_BODY =
    '<html><body>' +
    '<div id="divRplyFwdMsg">--- quoted ---</div>' +
    '</body></html>';

  function stubDraftFlow(client: ReturnType<typeof makeCtx>['client']) {
    vi.spyOn(client, 'post').mockImplementation(async (path: string) => {
      if (/create(Reply|ReplyAll|Forward)/i.test(path)) {
        return { id: 'DRAFT_HTML' } as Message as never;
      }
      return undefined as never; // send step
    });
    vi.spyOn(client, 'get').mockResolvedValue({
      id: 'DRAFT_HTML',
      body: { contentType: 'html', content: HTML_DRAFT_BODY },
    } as Message as never);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue(undefined as never);
    return patch;
  }

  it('reply with HTML sig uses createReply → PATCH → send (three POSTs/PATCH)', async () => {
    const { ctx, client, out } = makeCtx(['MSG1', '--body', 'Hi', '--json'], htmlEnv);
    const patch = stubDraftFlow(client);

    await replyCommand.handler(ctx);

    // POST createReply + POST send = 2 post calls; PATCH body = 1 patch call
    expect(patch).toHaveBeenCalledTimes(1);
    const patched = patch.mock.calls[0]?.[1] as { body: { content: string } };
    expect(patched.body.content).toContain('outlook-cli-signature');
    expect(patched.body.content).toContain('<b>Griffin</b>');
    expect(patched.body.content).toContain('Hi');
    const result = JSON.parse(text(out));
    expect(result).toMatchObject({ signatureApplied: true, htmlSignature: true });
  });

  it('reply-all with HTML sig uses createReplyAll path', async () => {
    const { ctx, client } = makeCtx(['MSG1'], htmlEnv);
    stubDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await replyAllCommand.handler(ctx);

    const createCall = postSpy.mock.calls.find(([p]) =>
      typeof p === 'string' && p.includes('createReplyAll'),
    );
    expect(createCall).toBeDefined();
  });

  it('forward with HTML sig uses createForward and passes toRecipients', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1', '--to', 'a@example.com', '--body', 'FYI'],
      htmlEnv,
    );
    stubDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await forwardCommand.handler(ctx);

    const createCall = postSpy.mock.calls.find(([p]) =>
      typeof p === 'string' && p.includes('createForward'),
    );
    expect(createCall).toBeDefined();
    expect(createCall?.[1]).toMatchObject({
      toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
    });
  });

  it('with both text+HTML sig configured, HTML flow is used and sig div is in patch', async () => {
    const { ctx, client } = makeCtx(['MSG1', '--body', 'Hey'], htmlTextEnv);
    const patch = stubDraftFlow(client);

    await replyCommand.handler(ctx);

    const patched = patch.mock.calls[0]?.[1] as { body: { content: string } };
    expect(patched.body.content).toContain('outlook-cli-signature');
  });

  it('reply --dry-run with HTML sig shows three-step output without network calls', async () => {
    const { ctx, fetchImpl, out } = makeCtx(['MSG1', '--body', 'Hi', '--dry-run'], htmlEnv);

    await replyCommand.handler(ctx);

    expect(fetchImpl).not.toHaveBeenCalled();
    expect(text(out)).toContain('DRY RUN');
    expect(text(out)).toContain('Step 1');
    expect(text(out)).toContain('Step 2');
    expect(text(out)).toContain('Step 3');
    expect(text(out)).toContain('createReply');
  });

  it('forward --dry-run with HTML sig shows all three steps', async () => {
    const { ctx, fetchImpl, out } = makeCtx(
      ['MSG1', '--to', 'a@example.com', '--dry-run'],
      htmlEnv,
    );

    await forwardCommand.handler(ctx);

    expect(fetchImpl).not.toHaveBeenCalled();
    expect(text(out)).toContain('createForward');
    expect(text(out)).toContain('Step 3');
  });
});
