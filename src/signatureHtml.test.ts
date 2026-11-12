/**
 * Unit tests for src/signatureHtml.ts
 *
 * Covers: body extraction, data: → cid: rewriting, deduplication,
 * unsupported MIME pass-through, and the plain-HTML no-op invariant.
 */

import { describe, expect, it } from 'vitest';
import {
  extractBodyHtml,
  processHtmlSignature,
  rewriteDataUris,
} from './signatureHtml.js';

// Minimal 1×1 transparent PNG as base64 (11 bytes decoded).
const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

const JPEG_B64 =
  '/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AJQAB/9k=';

describe('extractBodyHtml', () => {
  it('returns undefined for a plain HTML snippet with no <body>', () => {
    expect(extractBodyHtml('<b>hello</b>')).toBeUndefined();
  });

  it('extracts inner content between <body> and </body>', () => {
    const html = '<html><head><title>T</title></head><body><p>Hello</p></body></html>';
    expect(extractBodyHtml(html)).toBe('<p>Hello</p>');
  });

  it('handles <body> with attributes', () => {
    const html = '<html><body style="font-family:Arial"><p>Hi</p></body></html>';
    expect(extractBodyHtml(html)).toBe('<p>Hi</p>');
  });

  it('is case-insensitive on <BODY> and </BODY>', () => {
    const html = '<HTML><BODY><p>Hi</p></BODY></HTML>';
    expect(extractBodyHtml(html)).toBe('<p>Hi</p>');
  });

  it('uses the last </body> when multiple appear (malformed HTML)', () => {
    const html = '<body>first</body><body>second</body>';
    // last open <body> is the second one; last </body> is the second one
    expect(extractBodyHtml(html)).toBe('second');
  });

  it('returns content after <body> when no </body> is present', () => {
    const html = '<html><body><p>unclosed';
    expect(extractBodyHtml(html)).toBe('<p>unclosed');
  });

  it('strips leading whitespace after the <body> tag', () => {
    const html = '<body>\n  <p>Hi</p>\n</body>';
    // The regex strips the newline+spaces after <body> as part of `(?:\s*)`.
    expect(extractBodyHtml(html)).toBe('<p>Hi</p>\n');
  });
});

describe('rewriteDataUris', () => {
  it('replaces a data: PNG src with cid: and returns an inline attachment', () => {
    const html = `<img src="data:image/png;base64,${PNG_B64}">`;
    const { html: out, inlineAttachments } = rewriteDataUris(html);

    expect(out).toMatch(/src="cid:sig-img-[0-9a-f]{8}@outlook-cli"/);
    expect(out).not.toContain('data:');
    expect(inlineAttachments).toHaveLength(1);

    const att = inlineAttachments[0]!;
    expect(att['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(att.isInline).toBe(true);
    expect(att.contentType).toBe('image/png');
    expect(att.contentBytes).toBe(PNG_B64);
    expect(att.name).toBe('signature-image-1.png');
    expect(att.contentId).toMatch(/^sig-img-[0-9a-f]{8}@outlook-cli$/);
  });

  it('replaces a data: JPEG src and uses jpg extension', () => {
    const html = `<img src="data:image/jpeg;base64,${JPEG_B64}">`;
    const { inlineAttachments } = rewriteDataUris(html);
    expect(inlineAttachments[0]!.name).toMatch(/\.jpg$/);
    expect(inlineAttachments[0]!.contentType).toBe('image/jpeg');
  });

  it('deduplicates identical data: URIs — one attachment, two cid: refs', () => {
    const img = `<img src="data:image/png;base64,${PNG_B64}">`;
    const html = `${img}${img}`;
    const { html: out, inlineAttachments } = rewriteDataUris(html);

    expect(inlineAttachments).toHaveLength(1);
    // Both img tags should reference the same cid.
    const cidMatch = out.match(/cid:(sig-img-[0-9a-f]{8}@outlook-cli)/g);
    expect(cidMatch).toHaveLength(2);
    expect(cidMatch![0]).toBe(cidMatch![1]);
  });

  it('leaves unsupported data: MIME types untouched', () => {
    const html = `<object data="data:application/pdf;base64,AAAA">`;
    const { html: out, inlineAttachments } = rewriteDataUris(html);
    expect(out).toContain('data:application/pdf');
    expect(inlineAttachments).toHaveLength(0);
  });

  it('handles single-quoted src attributes', () => {
    const html = `<img src='data:image/png;base64,${PNG_B64}'>`;
    const { html: out, inlineAttachments } = rewriteDataUris(html);
    expect(out).toMatch(/src="cid:/);
    expect(inlineAttachments).toHaveLength(1);
  });

  it('produces deterministic content-ids (same input → same id)', () => {
    const html = `<img src="data:image/png;base64,${PNG_B64}">`;
    const r1 = rewriteDataUris(html);
    const r2 = rewriteDataUris(html);
    expect(r1.inlineAttachments[0]!.contentId).toBe(r2.inlineAttachments[0]!.contentId);
  });

  it('numbers multiple different images sequentially', () => {
    const html =
      `<img src="data:image/png;base64,${PNG_B64}">` +
      `<img src="data:image/jpeg;base64,${JPEG_B64}">`;
    const { inlineAttachments } = rewriteDataUris(html);
    expect(inlineAttachments).toHaveLength(2);
    expect(inlineAttachments[0]!.name).toBe('signature-image-1.png');
    expect(inlineAttachments[1]!.name).toBe('signature-image-2.jpg');
  });
});

describe('processHtmlSignature', () => {
  it('plain HTML snippet is returned byte-identical with no attachments', () => {
    const plain = '<b>Griffin Long</b><br>Agent AFK';
    const { html, inlineAttachments } = processHtmlSignature(plain);
    expect(html).toBe(plain);
    expect(inlineAttachments).toHaveLength(0);
  });

  it('extracts body from a full Outlook-style document', () => {
    const doc =
      '<html><head><style>body{font-family:Arial}</style></head>' +
      '<body><p>Sig content</p></body></html>';
    const { html, inlineAttachments } = processHtmlSignature(doc);
    expect(html).toBe('<p>Sig content</p>');
    expect(inlineAttachments).toHaveLength(0);
  });

  it('rewrites data: images in a full document and strips the head', () => {
    const doc =
      '<html><head><title>Sig</title></head>' +
      `<body><img src="data:image/png;base64,${PNG_B64}"><p>Hi</p></body></html>`;
    const { html, inlineAttachments } = processHtmlSignature(doc);

    expect(html).not.toContain('<html>');
    expect(html).not.toContain('<head>');
    expect(html).toContain('cid:');
    expect(html).not.toContain('data:');
    expect(inlineAttachments).toHaveLength(1);
  });

  it('rewrites data: images in a plain snippet (no <body>)', () => {
    const snippet = `<img src="data:image/png;base64,${PNG_B64}">`;
    const { html, inlineAttachments } = processHtmlSignature(snippet);
    expect(html).toMatch(/src="cid:/);
    expect(inlineAttachments).toHaveLength(1);
  });

  it('no-op invariant: empty inlineAttachments when no data: images present', () => {
    const doc = '<html><body><b>plain</b></body></html>';
    const { html, inlineAttachments } = processHtmlSignature(doc);
    expect(html).toBe('<b>plain</b>');
    expect(inlineAttachments).toHaveLength(0);
  });
});
