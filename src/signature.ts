/**
 * Email signatures.
 *
 * Microsoft Graph never applies the user's Outlook signature to mail sent
 * through the API (signatures are injected client-side by Outlook, and
 * mailboxSettings exposes no signature field), so the CLI appends one itself.
 *
 * Resolution, per invocation:
 *   --no-signature            -> none (explicit opt-out always wins)
 *   --signature <path>        -> that file (.html/.htm = HTML, else text)
 *   OUTLOOK_SIGNATURE_FILE    -> plain-text signature
 *   OUTLOOK_SIGNATURE_HTML_FILE -> HTML signature
 *   nothing configured        -> none
 *
 * Default-on once configured is deliberate: forgetting a flag must not mean
 * unsigned mail. Every surface reports `signatureApplied` in --json output and
 * --dry-run shows the appended text, so the behaviour is never invisible.
 *
 * A configured file that cannot be read is a hard error, not a silent skip:
 * sending unsigned mail when the user asked for a signature is the surprise.
 */

import { readFile } from 'node:fs/promises';
import type { ParsedArgs } from './args.js';
import { UsageError } from './errors.js';
import type { ItemBody } from './types.js';

export type EnvLike = Readonly<Record<string, string | undefined>>;

export const SIGNATURE_TEXT_ENV = 'OUTLOOK_SIGNATURE_FILE';
export const SIGNATURE_HTML_ENV = 'OUTLOOK_SIGNATURE_HTML_FILE';

/** RFC 3676 §4.3 signature separator ("-- " with the trailing space). */
export const TEXT_DELIMITER = '-- ';

export interface Signature {
  text?: string;
  html?: string;
}

function trimmed(value: string | undefined): string | undefined {
  const t = value?.trim();
  return t ? t : undefined;
}

function isHtmlPath(path: string): boolean {
  return /\.html?$/i.test(path);
}

async function readSignatureFile(path: string, origin: string): Promise<string> {
  try {
    const content = await readFile(path, 'utf8');
    // An all-whitespace file is a misconfiguration, not a valid empty sig.
    if (content.trim() === '') {
      throw new UsageError(
        `Signature file ${path} (from ${origin}) is empty. ` +
          'Add content, or pass --no-signature to send without one.',
      );
    }
    return content;
  } catch (e) {
    if (e instanceof UsageError) throw e;
    const reason = e instanceof Error ? e.message : String(e);
    throw new UsageError(
      `Could not read signature file ${path} (from ${origin}): ${reason}\n` +
        'Fix the path, or pass --no-signature to send without one.',
    );
  }
}

/**
 * Decide which signature (if any) applies. `--signature` and `--no-signature`
 * share the key `signature`: args.ts records `--no-x` as `x=false`.
 */
export async function resolveSignature(
  args: ParsedArgs,
  env: EnvLike = {},
): Promise<Signature | undefined> {
  if (args.has('signature')) {
    const path = args.str('signature');
    if (path === undefined) {
      // A string means a path; anything else is a boolean. `bool()` throws
      // on ambiguous values, which is what we want for `--signature=maybe`.
      if (!args.bool('signature')) return undefined; // --no-signature
      throw new UsageError(
        '--signature needs a file path (e.g. --signature ~/sig.txt). ' +
          'Use --no-signature to send without one.',
      );
    }
    if (path.trim() === '' || /^(false|0|no|off)$/i.test(path.trim())) {
      return undefined; // --signature=false and friends
    }
    const content = await readSignatureFile(path, '--signature');
    return isHtmlPath(path) ? { html: content } : { text: content };
  }

  const textPath = trimmed(env[SIGNATURE_TEXT_ENV]);
  const htmlPath = trimmed(env[SIGNATURE_HTML_ENV]);
  if (textPath === undefined && htmlPath === undefined) return undefined;

  const sig: Signature = {};
  if (textPath !== undefined) {
    sig.text = await readSignatureFile(textPath, SIGNATURE_TEXT_ENV);
  }
  if (htmlPath !== undefined) {
    sig.html = await readSignatureFile(htmlPath, SIGNATURE_HTML_ENV);
  }
  return sig;
}

/** True when the user asked FOR a signature via flag (path or bare --signature). */
export function signatureFlagRequested(args: ParsedArgs): boolean {
  if (!args.has('signature')) return false;
  const v = args.str('signature');
  if (v === undefined) return args.bool('signature');
  return !(v.trim() === '' || /^(false|0|no|off)$/i.test(v.trim()));
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Append a plain-text signature with the RFC 3676 delimiter. */
export function appendTextSignature(content: string, sig: Signature): string {
  if (sig.text === undefined) {
    throw new UsageError(
      'Only an HTML signature is configured, but this message body is plain ' +
        `text. Set ${SIGNATURE_TEXT_ENV}, pass --html (send/draft), ` +
        'or pass --no-signature.',
    );
  }
  const head = content.trimEnd();
  const tail = `${TEXT_DELIMITER}\n${sig.text.trimEnd()}`;
  return head === '' ? tail : `${head}\n\n${tail}`;
}

/**
 * Append an HTML signature, inside </body> when the body has one.
 *
 * Trust boundary: `sig.html` is inserted verbatim without sanitisation.
 * This is intentional — the file belongs to the user and may contain
 * arbitrary HTML (images, inline styles, etc.). Treat it like the token
 * cache: it should be writable only by its owner (`chmod 600`). Do not
 * point this at untrusted or world-writable paths.
 */
function appendHtmlSignature(content: string, sig: Signature): string {
  const html =
    sig.html?.trim() ??
    escapeHtml((sig.text ?? '').trimEnd()).replace(/\r?\n/g, '<br>');
  const block = `<br><div class="outlook-cli-signature">${html}</div>`;

  // Search the original string case-insensitively. Searching a lowercased
  // copy breaks when lowercasing changes length (U+0130 "İ" becomes two code
  // units), shifting the splice into the middle of the tag.
  const matches = [...content.matchAll(/<\/body>/gi)];
  const last = matches[matches.length - 1];
  if (last?.index === undefined) return `${content}${block}`;
  return `${content.slice(0, last.index)}${block}${content.slice(last.index)}`;
}

/** Return a new body with the signature appended in the body's own format. */
export function applySignature(body: ItemBody, sig: Signature): ItemBody {
  const content =
    body.contentType.toLowerCase() === 'html'
      ? appendHtmlSignature(body.content, sig)
      : appendTextSignature(body.content, sig);
  return { contentType: body.contentType, content };
}
