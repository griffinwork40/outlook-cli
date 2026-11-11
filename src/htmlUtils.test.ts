/**
 * Unit tests for escapeHtml and stripHtmlTags in htmlUtils.ts.
 */

import { describe, expect, it } from 'vitest';
import { escapeHtml, stripHtmlTags } from './htmlUtils.js';

describe('escapeHtml', () => {
  it('escapes &, <, >, ", and \'', () => {
    expect(escapeHtml('a & b')).toBe('a &amp; b');
    expect(escapeHtml('<tag>')).toBe('&lt;tag&gt;');
    expect(escapeHtml('"quoted"')).toBe('&quot;quoted&quot;');
    expect(escapeHtml("it's")).toBe('it&#39;s');
  });

  it('leaves plain text untouched', () => {
    expect(escapeHtml('hello world')).toBe('hello world');
  });
});

describe('stripHtmlTags', () => {
  it('strips a simple tag', () => {
    expect(stripHtmlTags('<b>bold</b>')).toBe('bold');
  });

  it('strips multiple tags', () => {
    expect(stripHtmlTags('<b>Alice</b> <a href="x">link</a>')).toBe('Alice link');
  });

  it('handles > inside a double-quoted attribute value', () => {
    // The old regex `/<[^>]*>/g` would break on this, stopping at the first `>`
    // inside the attribute and leaving `">` as literal text.
    expect(stripHtmlTags('<a title="a>b">click</a>')).toBe('click');
  });

  it('handles > inside a single-quoted attribute value', () => {
    expect(stripHtmlTags("<a title='x>y'>text</a>")).toBe('text');
  });

  it('strips HTML comments', () => {
    expect(stripHtmlTags('<!-- comment -->text')).toBe('text');
  });

  it('leaves text with no tags unchanged', () => {
    expect(stripHtmlTags('plain text')).toBe('plain text');
  });

  it('treats an unterminated tag as literal text', () => {
    // `<b` with no closing `>` — the opening `<` should be preserved as-is.
    expect(stripHtmlTags('hello <b world')).toBe('hello <b world');
  });

  it('handles self-closing tags', () => {
    expect(stripHtmlTags('line<br/>end')).toBe('lineend');
    expect(stripHtmlTags('<img src="x.png" />')).toBe('');
  });

  it('strips a realistic HTML signature', () => {
    const html =
      '<div><b>Alice Smith</b><br>' +
      '<a href="https://example.com?a=1&amp;b=2" title="a>b">example.com</a>' +
      '</div>';
    const result = stripHtmlTags(html);
    expect(result).toBe('Alice Smithexample.com');
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
  });
});
