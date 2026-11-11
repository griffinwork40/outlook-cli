/**
 * Outlook-exported HTML signature processing.
 *
 * Outlook saves signatures as full HTML documents with embedded images as
 * `data:` URIs.  Two transformations are needed before such a file can be
 * used as a mail signature:
 *
 *   1. Body extraction — use only the content of <body>…</body>, discarding
 *      <html>, <head>, and any <style> blocks in the head.  Styles in a
 *      nested <head> would not apply inside the composed message body anyway,
 *      and the outer document already has its own <head>.  Any <style> tags
 *      that appear INSIDE the <body> (unusual but possible) are preserved
 *      verbatim, since they may apply to signature elements.
 *
 *   2. data: image rewriting — each `src="data:image/(png|jpeg|gif|webp);
 *      base64,<data>"` is replaced with `src="cid:<contentId>"` and the
 *      decoded bytes are returned as an inline attachment.  The content-id is
 *      derived deterministically from a SHA-256 of the raw data URI string so
 *      identical images in different signatures share one attachment.
 *
 * Plain signatures (no <body>, no data: images) pass through unchanged, so
 * this module is a no-op for simple HTML strings.
 *
 * Trust boundary: input HTML is operator-supplied and treated as trusted,
 * matching the existing boundary in signature.ts.
 */

import { createHash } from 'node:crypto';

/** Subset of FileAttachmentUpload used for inline signature images. */
export interface InlineImageAttachment {
  '@odata.type': '#microsoft.graph.fileAttachment';
  name: string;
  contentType: string;
  contentBytes: string; // base64
  isInline: true;
  contentId: string;
}

export interface ProcessedHtml {
  /** The signature HTML ready to embed (body-only when input was a full doc). */
  html: string;
  /** Inline image attachments to add alongside the message, may be empty. */
  inlineAttachments: InlineImageAttachment[];
}

/** MIME types we handle; covers every format Outlook embeds. */
const SUPPORTED_IMAGE_MIME = new Set([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/webp',
]);

/**
 * Extension map — used to build a readable filename for the attachment.
 * Falls back to the sub-type (e.g. "jpeg") for unknown MIME types.
 */
const MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
};

/**
 * Extract <body> inner HTML from a full HTML document.
 * Returns undefined when no <body> tag is found.
 */
export function extractBodyHtml(html: string): string | undefined {
  // Find the last <body ...> open tag (case-insensitive).
  const openMatches = [...html.matchAll(/<body(?:\s[^>]*)?>(?:\s*)/gi)];
  if (openMatches.length === 0) return undefined;
  const openMatch = openMatches[openMatches.length - 1]!;
  const start = openMatch.index! + openMatch[0].length;

  // Find the matching </body>.
  const closeMatches = [...html.matchAll(/<\/body>/gi)];
  if (closeMatches.length === 0) return html.slice(start);
  const closeMatch = closeMatches[closeMatches.length - 1]!;
  return html.slice(start, closeMatch.index);
}

/**
 * Derive a stable content-id from the data URI (SHA-256 of the first 256
 * chars of the base64 data, which is plenty for uniqueness in practice).
 * Format: `sig-img-<hex8>@outlook-cli`.
 */
function contentIdFromData(dataUri: string): string {
  const sample = dataUri.slice(0, 256);
  const hex = createHash('sha256').update(sample).digest('hex').slice(0, 8);
  return `sig-img-${hex}@outlook-cli`;
}

/**
 * Rewrite data: image URIs to cid: references and collect inline attachments.
 *
 * We match the full src="data:image/...;base64,..." attribute (single or
 * double quotes).  Each unique data URI produces one attachment; duplicates
 * reuse the same content-id so the attachment is not added twice.
 */
export function rewriteDataUris(html: string): {
  html: string;
  inlineAttachments: InlineImageAttachment[];
} {
  // Map from contentId → attachment (deduplicates identical images).
  const seen = new Map<string, InlineImageAttachment>();
  const order: string[] = []; // preserve insertion order

  // Match src="data:image/TYPE;base64,DATA" (single or double quotes).
  // The base64 data runs until the closing quote.
  const result = html.replace(
    /src=(["'])data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=]+)\1/g,
    (_match, _quote, mimeType: string, base64Data: string) => {
      const mime = mimeType.toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        // Leave unsupported types untouched.
        return _match;
      }

      const dataUri = `data:${mime};base64,${base64Data}`;
      const contentId = contentIdFromData(dataUri);

      if (!seen.has(contentId)) {
        const ext = MIME_EXT[mime] ?? mime.split('/')[1] ?? 'bin';
        const name = `signature-image-${order.length + 1}.${ext}`;
        const attachment: InlineImageAttachment = {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name,
          contentType: mime,
          contentBytes: base64Data,
          isInline: true,
          contentId,
        };
        seen.set(contentId, attachment);
        order.push(contentId);
      }

      return `src="cid:${contentId}"`;
    },
  );

  return {
    html: result,
    inlineAttachments: order.map((id) => seen.get(id)!),
  };
}

/**
 * Process an HTML signature file, applying body extraction and data: → cid:
 * rewriting.  Returns the transformed HTML and any inline attachments.
 *
 * A plain HTML snippet (no <body>, no data: images) is returned unchanged
 * with an empty attachments array — byte-identical to what signature.ts
 * produced before this module existed.
 */
export function processHtmlSignature(raw: string): ProcessedHtml {
  // Step 1: extract body content when this is a full document.
  const bodyContent = extractBodyHtml(raw);
  const html = bodyContent !== undefined ? bodyContent : raw;

  // Step 2: rewrite data: URIs → cid: references.
  const { html: rewritten, inlineAttachments } = rewriteDataUris(html);

  return { html: rewritten, inlineAttachments };
}
