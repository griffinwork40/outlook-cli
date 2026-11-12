/**
 * Additional insertion-behaviour tests for executeReplyDraft:
 *  - quote-aware HTML tag stripping (> inside quoted attribute)
 *  - first-match <body> injection (not last)
 *
 * Extracted into a separate file to stay within the per-file LOC ceiling.
 */

import { describe, expect, it, vi } from 'vitest';
import { GraphClient } from '../client.js';
import type { Message } from '../types.js';
import { executeReplyDraft } from './replyDraft.js';
import { makeClient } from './testUtils.js';

// ---------------------------------------------------------------------------
// HTML-only sig in text draft: quote-aware tag stripping
// ---------------------------------------------------------------------------
describe('executeReplyDraft — text draft, html-only sig tag stripping', () => {
  function makeTextClient(draftBody: string) {
    const { client } = makeClient();
    vi.spyOn(client, 'post').mockImplementation(async (path: string) => {
      if (path.includes('/createReplyAll')) return { id: 'DRAFT_STRIP' } as Message as never;
      return undefined as never;
    });
    vi.spyOn(client, 'get').mockResolvedValue({
      id: 'DRAFT_STRIP',
      body: { contentType: 'text', content: draftBody },
    } as Message as never);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue(undefined as never);
    return { client, patch };
  }

  it('handles > inside a double-quoted attribute value without leaving tag fragments', async () => {
    // `<a title="a>b">` has a `>` inside an attribute value.
    // The old regex `/<[^>]*>/g` would stop at the inner `>` and leave `">` as text.
    const { client, patch } = makeTextClient('original');

    await executeReplyDraft(client, {
      messageId: 'MSG2',
      action: 'createReplyAll',
      userText: '',
      signature: { html: '<a title="a>b">Alice</a>' },
    });

    const patched = patch.mock.calls[0]?.[1] as { body: { content: string } };
    const content = patched.body.content;
    expect(content).toContain('Alice');
    expect(content).not.toContain('<a');
    expect(content).not.toMatch(/^">/m);
  });
});

// ---------------------------------------------------------------------------
// HTML draft: first <body> tag is used for injection, not the last
// ---------------------------------------------------------------------------
describe('executeReplyDraft — <body> injection uses first match', () => {
  it('injects after the first <body> tag, not the last', async () => {
    // Draft with two <body> occurrences: the real one, plus one inside a comment
    // in the quoted section.
    const twoBody =
      '<html><body>' +
      '<div><!-- quoted: <body class="quoted"> --></div>' +
      '</body></html>';
    const { client } = makeClient();
    vi.spyOn(client, 'post').mockImplementation(async (path: string) => {
      if (path.includes('/createReply')) return { id: 'DRAFT_BODY' } as Message as never;
      return undefined as never;
    });
    vi.spyOn(client, 'get').mockResolvedValue({
      id: 'DRAFT_BODY',
      body: { contentType: 'html', content: twoBody },
    } as Message as never);
    const patch = vi.spyOn(client, 'patch').mockResolvedValue(undefined as never);

    await executeReplyDraft(client, {
      messageId: 'MSG1',
      action: 'createReply',
      userText: 'hi',
      signature: { html: '<b>Sig</b>' },
    });

    const patched = patch.mock.calls[0]?.[1] as { body: { content: string } };
    const content = patched.body.content;

    // Insertion block starts immediately after the first <body> tag.
    const firstBodyMatch = /<body[^>]*>/i.exec(content);
    expect(firstBodyMatch).not.toBeNull();
    const afterFirstBody = firstBodyMatch!.index + firstBodyMatch![0].length;
    expect(content[afterFirstBody]).toBe('<');

    // Sig appears between first <body> and </body>.
    const sigIdx = content.indexOf('outlook-cli-signature');
    expect(sigIdx).toBeGreaterThan(afterFirstBody);
    expect(sigIdx).toBeLessThan(content.indexOf('</body>'));

    // Sig appears before the second <body> occurrence (inside the comment).
    const secondBodyIdx = content.indexOf('<body', afterFirstBody + 1);
    if (secondBodyIdx !== -1) {
      expect(sigIdx).toBeLessThan(secondBodyIdx);
    }
  });
});
