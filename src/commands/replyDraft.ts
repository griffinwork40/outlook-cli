/**
 * Three-step reply/replyAll/forward flow used when an HTML signature applies.
 *
 * The one-shot `/reply` action only accepts a `comment` field (plain text),
 * so it cannot carry an HTML signature.  When the resolved signature has an
 * HTML form we fall back to:
 *
 *   1. POST /me/messages/{id}/createReply   → draft (Exchange adds quoted body)
 *   2. PATCH /me/messages/{draftId}         → insert user text + signature
 *   3. POST /me/messages/{draftId}/send     → send
 *
 * Threading headers (In-Reply-To, References) are set by step 1 and survive
 * the PATCH, so conversation threading is preserved.
 *
 * Signature trust boundary: sig.html is operator-supplied (--signature <path>
 * or OUTLOOK_SIGNATURE_HTML_FILE) and is treated as trusted content, the same
 * boundary documented in signature.ts and README.md §Signature files.
 * Never point these variables at world-writable or user-uploaded paths.
 *
 * Atomicity: the flow is NOT atomic.  On PATCH or send failure we attempt to
 * DELETE the draft; if that also fails, the user is instructed to run
 * `outlook draft list` to locate and delete it manually.
 *
 * docs/02 §8: POST is never auto-retried; PATCH/DELETE use the normal retry
 * logic already in GraphClient.
 */

import type { GraphClient } from '../client.js';
import { OutlookCliError } from '../errors.js';
import {
  TEXT_DELIMITER,
  type Signature,
} from '../signature.js';
import type { ItemBody, Message } from '../types.js';

export type CreateAction = 'createReply' | 'createReplyAll' | 'createForward';

export interface ReplyDraftOptions {
  messageId: string;
  action: CreateAction;
  /** User-supplied body text (may be empty string). */
  userText: string;
  signature: Signature;
  /** For forward: required recipients. */
  toRecipients?: Array<{ emailAddress: { address: string; name?: string } }>;
}

export interface ReplyDraftSpecs {
  createSpec: ReturnType<GraphClient['buildSpec']>;
  patchSpec: ReturnType<GraphClient['buildSpec']>;
  sendSpec: ReturnType<GraphClient['buildSpec']>;
}

/**
 * Build the three RequestSpecs for --dry-run display.
 *
 * The PATCH body cannot be fully materialised at dry-run time (we don't know
 * the draft id or its Exchange-generated body), so we produce a representative
 * placeholder that shows the shape of what will be sent.
 */
export function buildReplyDraftSpecs(
  client: GraphClient,
  opts: ReplyDraftOptions,
): ReplyDraftSpecs {
  const basePath = `/me/messages/${encodeURIComponent(opts.messageId)}`;

  const createBody: Record<string, unknown> = {};
  if (opts.toRecipients) createBody['toRecipients'] = opts.toRecipients;

  const createSpec = client.buildSpec('POST', `${basePath}/${opts.action}`, createBody);

  // Representative patch body — actual content depends on the draft Exchange returns.
  const patchPreview = buildPatchBody(
    { contentType: 'html', content: '<html><body><!-- Exchange quoted body --></body></html>' },
    opts.userText,
    opts.signature,
  );
  const patchSpec = client.buildSpec('PATCH', '/me/messages/{draftId}', patchPreview);

  const sendSpec = client.buildSpec('POST', '/me/messages/{draftId}/send', undefined);

  return { createSpec, patchSpec, sendSpec };
}

/**
 * Execute the three-step flow, cleaning up the draft on failure.
 *
 * Returns the draft id so callers can surface it in the success output.
 */
export async function executeReplyDraft(
  client: GraphClient,
  opts: ReplyDraftOptions,
): Promise<string> {
  const basePath = `/me/messages/${encodeURIComponent(opts.messageId)}`;

  // Step 1: create draft (Exchange fills in quoted body + threading headers).
  const createBody: Record<string, unknown> = {};
  if (opts.toRecipients) createBody['toRecipients'] = opts.toRecipients;

  const draft = await client.post<Message>(
    `${basePath}/${opts.action}`,
    createBody,
  );
  const draftId = draft?.id;
  if (!draftId) {
    throw new OutlookCliError(
      `${opts.action} did not return a draft id. ` +
        'The message may not have been sent.',
    );
  }

  // Steps 2+3: patch then send, deleting the draft on any failure.
  try {
    // Step 2: read back the draft body Exchange generated, then patch.
    const draftFull = await client.get<Message>(
      `/me/messages/${encodeURIComponent(draftId)}`,
      { $select: 'id,body' },
    );
    if (!draftFull?.body) {
      throw new OutlookCliError(
        'Exchange returned a draft with no body. Cannot patch reply.',
      );
    }
    const draftBody: ItemBody = draftFull.body;
    const patchBody = buildPatchBody(draftBody, opts.userText, opts.signature);
    await client.patch<unknown>(
      `/me/messages/${encodeURIComponent(draftId)}`,
      patchBody,
    );

    // Step 3: send the patched draft.
    await client.post<void>(
      `/me/messages/${encodeURIComponent(draftId)}/send`,
      undefined,
    );
  } catch (stepErr) {
    // Attempt cleanup — best effort.
    let cleanupErr: unknown;
    try {
      await client.del(`/me/messages/${encodeURIComponent(draftId)}`);
    } catch (e) {
      cleanupErr = e;
    }

    const base =
      stepErr instanceof Error ? stepErr.message : String(stepErr);

    if (cleanupErr !== undefined) {
      const cleanMsg =
        cleanupErr instanceof Error ? cleanupErr.message : String(cleanupErr);
      throw new OutlookCliError(
        `Failed to send reply (${base}). ` +
          `Cleanup also failed (${cleanMsg}). ` +
          'A draft was left behind (ORPHAN) — run `outlook draft list` to find and delete it.',
      );
    }

    throw new OutlookCliError(
      `Failed to send reply (${base}). ` +
        'The draft was deleted. No message was sent.',
    );
  }

  return draftId;
}

// ---------------------------------------------------------------------------
// Body-insertion helpers
// ---------------------------------------------------------------------------

/**
 * Produce the PATCH body that inserts `userText` + signature above the
 * Exchange-generated quoted block, in the draft's own content type.
 */
function buildPatchBody(
  draftBody: ItemBody,
  userText: string,
  sig: Signature,
): { body: ItemBody } {
  const isHtml = draftBody.contentType.toLowerCase() === 'html';
  const newBody = isHtml
    ? insertHtml(draftBody.content, userText, sig)
    : insertText(draftBody.content, userText, sig);
  return { body: { contentType: draftBody.contentType, content: newBody } };
}

/**
 * Insert user text and HTML signature above the Exchange quote in an HTML body.
 *
 * Exchange wraps the quoted message in a <div id="divRplyFwdMsg"> block.  We
 * look for that marker and inject above it; if not found we prepend to <body>
 * (or to the whole content when there is no <body> tag).
 */
function insertHtml(draftContent: string, userText: string, sig: Signature): string {
  const textHtml = escapeHtml(userText.trimEnd()).replace(/\r?\n/g, '<br>');
  const sigHtml =
    sig.html?.trim() ??
    escapeHtml((sig.text ?? '').trimEnd()).replace(/\r?\n/g, '<br>');

  const insertion =
    `<div>${textHtml}</div>` +
    `<br><div class="outlook-cli-signature">${sigHtml}</div><br>`;

  // Try to place before the Exchange quote div.
  const quoteMatch = draftContent.match(/<div\s+id=["']divRplyFwdMsg["']/i);
  if (quoteMatch?.index !== undefined) {
    return (
      draftContent.slice(0, quoteMatch.index) +
      insertion +
      draftContent.slice(quoteMatch.index)
    );
  }

  // Fall back: inject at the start of <body>.
  const bodyMatch = [...draftContent.matchAll(/<body[^>]*>/gi)];
  const lastBody = bodyMatch[bodyMatch.length - 1];
  if (lastBody?.index !== undefined) {
    const afterBody = lastBody.index + lastBody[0].length;
    return (
      draftContent.slice(0, afterBody) +
      insertion +
      draftContent.slice(afterBody)
    );
  }

  // No <body> tag — prepend.
  return insertion + draftContent;
}

/**
 * Insert user text and text signature above the Exchange plain-text quote.
 *
 * Exchange delimits the quote with "-----Original Message-----" (or a locale
 * variant).  We look for that marker and insert above it; if not found we
 * prepend the new content.
 */
function insertText(draftContent: string, userText: string, sig: Signature): string {
  if (sig.text === undefined) {
    // HTML-only signature in a text draft: fall back to escaping the HTML.
    const stripped = (sig.html ?? '').replace(/<[^>]*>/g, '').trim();
    return buildTextInsertion(draftContent, userText, stripped);
  }
  return buildTextInsertion(draftContent, userText, sig.text.trimEnd());
}

function buildTextInsertion(
  draftContent: string,
  userText: string,
  sigText: string,
): string {
  const head = userText.trimEnd();
  const sigBlock = `${TEXT_DELIMITER}\n${sigText}`;
  const newContent = head === '' ? sigBlock : `${head}\n\n${sigBlock}`;

  // Look for Exchange's plain-text quote separator.
  // Use a structural pattern (---…text…---) rather than English "Original Message"
  // so localized variants (German, French, Japanese, etc.) are matched too.
  const sepMatch = draftContent.match(/^-{3,}[^-\r\n][^\r\n]*-{3,}/im);
  if (sepMatch?.index !== undefined) {
    return `${newContent}\n\n${draftContent.slice(sepMatch.index)}`;
  }

  // No quote found — prepend.
  if (draftContent.trim() === '') return newContent;
  return `${newContent}\n\n${draftContent}`;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
