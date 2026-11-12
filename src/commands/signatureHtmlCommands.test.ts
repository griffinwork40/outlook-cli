/**
 * Integration tests for Outlook-exported HTML signature handling at each
 * command surface: send, draft create, reply/reply-all/forward.
 *
 * Covers:
 *  - body extraction + data: → cid: rewriting wired into resolveSignature
 *  - inline attachments present in sendMail payload (send)
 *  - inline attachments present in draft create payload
 *  - inline attachments POSTed in the reply three-step flow
 *  - dry-run lists inline attachments without base64
 *  - plain-HTML signatures without images behave byte-identically to before
 */

import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { Message } from '../types.js';
import { draftCommand } from './draft.js';
import { forwardCommand, replyAllCommand, replyCommand } from './reply.js';
import { sendCommand } from './send.js';
import { makeCtx, text } from './testUtils.js';

// Minimal 1×1 transparent PNG as base64.
const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

const SEND_ARGS = ['--to', 'a@example.com', '--subject', 'S', '--body', 'Hello', '--html'];

let outlookSigPath: string; // full HTML doc with embedded PNG
let plainHtmlSigPath: string; // simple <b>Griffin</b> — no body, no data:

beforeAll(async () => {
  const dir = await mkdtemp(join(tmpdir(), 'outlook-htmlsig-cmd-'));

  // Simulate an Outlook-exported .htm file: full document with embedded PNG.
  const outlookHtm =
    '<html><head><style>body{font-family:Arial}</style></head>' +
    `<body><p>Griffin Long</p><img src="data:image/png;base64,${PNG_B64}"></body></html>`;
  outlookSigPath = join(dir, 'outlook.htm');
  await writeFile(outlookSigPath, outlookHtm);

  // Plain HTML snippet — no <body>, no data: images.
  plainHtmlSigPath = join(dir, 'plain.html');
  await writeFile(plainHtmlSigPath, '<b>Griffin</b>');
});

// ---------------------------------------------------------------------------
// send
// ---------------------------------------------------------------------------

describe('send with Outlook-exported HTML signature', () => {
  it('merges inline attachment into sendMail payload', async () => {
    const { ctx, client } = makeCtx(
      [...SEND_ARGS, '--signature', outlookSigPath],
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: {
        body: { content: string };
        attachments: Array<{ isInline: boolean; contentId: string; contentBytes: string }>;
      };
    };

    // Body must not contain data: URIs.
    expect(payload.message.body.content).not.toContain('data:');
    // Body must contain cid: reference.
    expect(payload.message.body.content).toContain('cid:');
    // Head stripped — no <html> or <head> in body content.
    expect(payload.message.body.content).not.toContain('<html>');
    expect(payload.message.body.content).not.toContain('<head>');

    // One inline attachment present.
    expect(payload.message.attachments).toHaveLength(1);
    const att = payload.message.attachments[0]!;
    expect(att.isInline).toBe(true);
    expect(att.contentId).toMatch(/^sig-img-[0-9a-f]{8}@outlook-cli$/);
    expect(att.contentBytes).toBe(PNG_B64);
  });

  it('dry-run does not print base64 contentBytes', async () => {
    const { ctx, out } = makeCtx(
      [...SEND_ARGS, '--signature', outlookSigPath, '--dry-run'],
    );

    await sendCommand.handler(ctx);

    const output = text(out);
    expect(output).toContain('DRY RUN');
    expect(output).not.toContain(PNG_B64);
    // The key should appear but its value redacted.
    expect(output).toContain('contentBytes');
    expect(output).toContain('[REDACTED]');
  });

  it('plain HTML signature without images is byte-identical to old behaviour', async () => {
    const { ctx, client } = makeCtx(
      [...SEND_ARGS, '--signature', plainHtmlSigPath],
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { body: { content: string }; attachments?: unknown[] };
    };
    // Exact content produced by the old path.
    expect(payload.message.body.content).toContain('<b>Griffin</b>');
    // No inline attachments added.
    expect(payload.message.attachments ?? []).toHaveLength(0);
  });

  it('via OUTLOOK_SIGNATURE_HTML_FILE env var also processes the signature', async () => {
    const { ctx, client } = makeCtx(
      SEND_ARGS,
      { OUTLOOK_SIGNATURE_HTML_FILE: outlookSigPath },
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { attachments: Array<{ isInline: boolean }> };
    };
    expect(payload.message.attachments).toHaveLength(1);
    expect(payload.message.attachments[0]!.isInline).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// draft create
// ---------------------------------------------------------------------------

describe('draft create with Outlook-exported HTML signature', () => {
  const DRAFT_MOCK: Message = { id: 'D1', isDraft: true };

  it('merges inline attachment into draft create payload', async () => {
    const { ctx, client } = makeCtx(
      ['create', '--subject', 'S', '--body', 'B', '--html', '--signature', outlookSigPath],
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(DRAFT_MOCK as never);

    await draftCommand.handler(ctx);

    const message = post.mock.calls[0]?.[1] as {
      body: { content: string };
      attachments: Array<{ isInline: boolean; contentId: string }>;
    };

    expect(message.body.content).toContain('cid:');
    expect(message.body.content).not.toContain('data:');
    expect(message.attachments).toHaveLength(1);
    expect(message.attachments[0]!.isInline).toBe(true);
  });

  it('dry-run redacts contentBytes in draft create output', async () => {
    const { ctx, out } = makeCtx(
      ['create', '--subject', 'S', '--body', 'B', '--html',
        '--signature', outlookSigPath, '--dry-run'],
    );

    await draftCommand.handler(ctx);

    const output = text(out);
    expect(output).toContain('DRY RUN');
    expect(output).not.toContain(PNG_B64);
  });
});

// ---------------------------------------------------------------------------
// reply three-step flow
// ---------------------------------------------------------------------------

const HTML_DRAFT_BODY =
  '<html><body><div id="divRplyFwdMsg">--- quoted ---</div></body></html>';

function stubHtmlDraftFlow(client: ReturnType<typeof makeCtx>['client']) {
  vi.spyOn(client, 'post').mockImplementation(async (path: string) => {
    if (/create(Reply|ReplyAll|Forward)/i.test(path)) {
      return { id: 'DRAFT_CID' } as Message as never;
    }
    return undefined as never; // send step + attachment POST
  });
  vi.spyOn(client, 'get').mockResolvedValue({
    id: 'DRAFT_CID',
    body: { contentType: 'html', content: HTML_DRAFT_BODY },
  } as Message as never);
  const patch = vi.spyOn(client, 'patch').mockResolvedValue(undefined as never);
  return patch;
}

describe('reply three-step flow with Outlook-exported HTML signature', () => {
  it('reply POSTs inline attachment to draft after PATCH', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1', '--body', 'Hi'],
      { OUTLOOK_SIGNATURE_HTML_FILE: outlookSigPath },
    );
    stubHtmlDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await replyCommand.handler(ctx);

    // Check that an attachment POST was made to /attachments endpoint.
    const attachPost = postSpy.mock.calls.find(
      ([p]) => typeof p === 'string' && p.includes('/attachments'),
    );
    expect(attachPost).toBeDefined();
    const attachPayload = attachPost?.[1] as { isInline: boolean; contentId: string };
    expect(attachPayload.isInline).toBe(true);
    expect(attachPayload.contentId).toMatch(/^sig-img-/);
  });

  it('reply-all also POSTs inline attachment', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1'],
      { OUTLOOK_SIGNATURE_HTML_FILE: outlookSigPath },
    );
    stubHtmlDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await replyAllCommand.handler(ctx);

    const attachPost = postSpy.mock.calls.find(
      ([p]) => typeof p === 'string' && p.includes('/attachments'),
    );
    expect(attachPost).toBeDefined();
  });

  it('forward also POSTs inline attachment', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1', '--to', 'x@example.com'],
      { OUTLOOK_SIGNATURE_HTML_FILE: outlookSigPath },
    );
    stubHtmlDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await forwardCommand.handler(ctx);

    const attachPost = postSpy.mock.calls.find(
      ([p]) => typeof p === 'string' && p.includes('/attachments'),
    );
    expect(attachPost).toBeDefined();
  });

  it('reply --dry-run lists inline images without base64', async () => {
    const { ctx, out, fetchImpl } = makeCtx(
      ['MSG1', '--body', 'Hi', '--dry-run'],
      { OUTLOOK_SIGNATURE_HTML_FILE: outlookSigPath },
    );

    await replyCommand.handler(ctx);

    expect(fetchImpl).not.toHaveBeenCalled();
    const output = text(out);
    expect(output).toContain('DRY RUN');
    expect(output).toContain('Step 2b');
    expect(output).toContain('inline signature image');
    expect(output).toContain('cid:sig-img-');
    expect(output).not.toContain(PNG_B64);
  });

  it('reply with plain HTML sig (no images) does not POST to /attachments', async () => {
    const { ctx, client } = makeCtx(
      ['MSG1', '--body', 'Hi'],
      { OUTLOOK_SIGNATURE_HTML_FILE: plainHtmlSigPath },
    );
    stubHtmlDraftFlow(client);
    const postSpy = vi.spyOn(client, 'post');

    await replyCommand.handler(ctx);

    const attachPost = postSpy.mock.calls.find(
      ([p]) => typeof p === 'string' && p.includes('/attachments'),
    );
    expect(attachPost).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Regression: plain-HTML signatures untouched
// ---------------------------------------------------------------------------

describe('regression: plain-HTML signatures (no <body>, no data:) are byte-identical', () => {
  it('send: body content is exactly what applySignature would have produced', async () => {
    const { ctx, client } = makeCtx(
      [...SEND_ARGS, '--signature', plainHtmlSigPath],
    );
    const post = vi.spyOn(client, 'post').mockResolvedValue(undefined as never);

    await sendCommand.handler(ctx);

    const payload = post.mock.calls[0]?.[1] as {
      message: { body: { content: string } };
    };
    // The body is HTML (--html flag), so signature is appended via appendHtmlSignature.
    // Expected: original body + <br><div class="outlook-cli-signature"><b>Griffin</b></div>
    expect(payload.message.body.content).toBe(
      'Hello<br><div class="outlook-cli-signature"><b>Griffin</b></div>',
    );
  });
});
