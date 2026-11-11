/**
 * Shared HTML utility helpers used by signature.ts and replyDraft.ts.
 */

/** Escape special HTML characters in plain text. */
export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Strip HTML tags from a string, respecting attribute quoting so that a `>`
 * inside a quoted attribute value (e.g. `<a title="a>b">`) does not
 * prematurely end the tag.
 *
 * The tokenizer handles:
 *   - double-quoted attribute values  "..."
 *   - single-quoted attribute values  '...'
 *   - unquoted attribute values (no embedded `>`)
 *   - comments <!-- ... -->
 *
 * This is intentionally minimal — it covers the HTML that appears in
 * user-supplied signature files, not arbitrary adversarial markup.
 */
export function stripHtmlTags(html: string): string {
  let result = '';
  let i = 0;
  const len = html.length;

  while (i < len) {
    if (html[i] !== '<') {
      result += html[i++];
      continue;
    }

    // We're at the start of a potential tag.
    const start = i;
    i++; // skip '<'

    // HTML comment: <!-- ... -->
    if (html.startsWith('!--', i)) {
      i += 3;
      const end = html.indexOf('-->', i);
      i = end === -1 ? len : end + 3;
      continue;
    }

    // Scan through the tag, respecting quoted attribute values.
    let closed = false;
    while (i < len) {
      const ch = html[i];
      if (ch === '>') {
        i++;
        closed = true;
        break;
      }
      if (ch === '"') {
        i++;
        while (i < len && html[i] !== '"') i++;
        if (i < len) i++; // closing quote
      } else if (ch === "'") {
        i++;
        while (i < len && html[i] !== "'") i++;
        if (i < len) i++; // closing quote
      } else {
        i++;
      }
    }

    if (!closed) {
      // Unterminated tag — treat everything from '<' to end of string as
      // literal text (we consumed it in the scan loop, so copy it back).
      result += html.slice(start);
    }
    // Otherwise the tag is consumed and produces no output.
  }

  return result;
}
