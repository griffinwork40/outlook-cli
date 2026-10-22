/**
 * Shared composition helpers for send / reply / forward / draft.
 *
 * Recipient parsing, body building and attachment loading live here exactly
 * once so the four compose surfaces cannot drift apart — a recipient bug
 * fixed in `send` must not still be live in `forward`.
 */

import { readFile } from 'node:fs/promises';
import type { ParsedArgs } from '../args.js';
import { UsageError } from '../errors.js';
import { readFileAttachment } from '../mime.js';
import { applySignature, type Signature } from '../signature.js';
import type {
  BodyContentType,
  FileAttachmentUpload,
  ItemBody,
  Recipient,
} from '../types.js';

/**
 * Accepts both repeated flags (`--to a --to b`) and comma-separated values
 * (`--to "a,b"`), because both are things people actually type.
 */
export function parseRecipients(
  values: readonly string[],
  flag: string,
): Recipient[] {
  const out: Recipient[] = [];
  for (const value of values) {
    for (const part of value.split(',')) {
      const address = part.trim();
      if (address.length === 0) continue;
      if (!address.includes('@')) {
        throw new UsageError(
          `--${flag} received "${address}", which is not an email address.`,
        );
      }
      out.push({ emailAddress: { address } });
    }
  }
  return out;
}

export function recipientsFromArgs(args: ParsedArgs, flag: string): Recipient[] {
  return parseRecipients(args.strAll(flag), flag);
}

/** `--body <text>` or `--body-file <path>`, never both. */
export async function buildBody(
  args: ParsedArgs,
): Promise<ItemBody | undefined> {
  const inline = args.str('body');
  const file = args.str('body-file');

  if (inline !== undefined && file !== undefined) {
    throw new UsageError('Use either --body or --body-file, not both.');
  }

  const content = file !== undefined ? await readFile(file, 'utf8') : inline;
  if (content === undefined) return undefined;

  // Graph accepts either casing; we author lowercase (docs/02).
  const contentType: BodyContentType = args.bool('html') ? 'html' : 'text';
  return { contentType, content };
}

/** Each --attach is read and base64-encoded. Over 3 MB throws (mime.ts). */
export async function loadAttachments(
  args: ParsedArgs,
): Promise<FileAttachmentUpload[]> {
  const out: FileAttachmentUpload[] = [];
  for (const path of args.strAll('attach')) {
    out.push(await readFileAttachment(path));
  }
  return out;
}

export interface ComposedMessage {
  subject?: string;
  body?: ItemBody;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  attachments?: FileAttachmentUpload[];
}

/** Build the `message` object shared by sendMail and draft create/update. */
export async function composeMessage(
  args: ParsedArgs,
): Promise<ComposedMessage> {
  const message: ComposedMessage = {};

  const subject = args.str('subject');
  if (subject !== undefined) message.subject = subject;

  const body = await buildBody(args);
  if (body !== undefined) message.body = body;

  const to = recipientsFromArgs(args, 'to');
  if (to.length > 0) message.toRecipients = to;

  const cc = recipientsFromArgs(args, 'cc');
  if (cc.length > 0) message.ccRecipients = cc;

  const bcc = recipientsFromArgs(args, 'bcc');
  if (bcc.length > 0) message.bccRecipients = bcc;

  const attachments = await loadAttachments(args);
  if (attachments.length > 0) message.attachments = attachments;

  return message;
}

/**
 * Apply a resolved signature to an already-composed message (mutates).
 *
 * Extracted from composeMessage so callers can run --to/--subject validation
 * and the 'Nothing to create' check BEFORE the signature file is even touched.
 *
 * Precondition: `args` must be the same ParsedArgs object that was used to
 * build `message` (i.e. `ctx.args`). When the message has no body yet, the
 * body's content-type is derived from `args.bool('html')`; if `args` were a
 * different object the format could disagree with the rest of the message.
 * Every current caller satisfies this by passing `ctx.args` directly.
 *
 * The `Object.keys(message).length === 0` guard exists for `draft create`:
 * that command checks for an empty message and exits with "Nothing to create"
 * before reaching the signature step. Returning early here preserves that
 * check — `send` never reaches this branch because it validates --to and
 * --subject first.
 */
export function signMessage(
  message: ComposedMessage,
  signature: Signature | undefined,
  args: ParsedArgs,
): void {
  if (signature === undefined || Object.keys(message).length === 0) return;
  const base: ItemBody = message.body ?? {
    contentType: args.bool('html') ? 'html' : 'text',
    content: '',
  };
  message.body = applySignature(base, signature);
}

/**
 * Default true. `--no-save-sent` is parsed by args.ts into `save-sent=false`,
 * so we probe the de-prefixed name.
 */
export function saveToSentItems(args: ParsedArgs): boolean {
  return args.has('save-sent') ? args.bool('save-sent') : true;
}

/** Summarize recipients for the human-readable confirmation line. */
export function describeRecipients(recipients?: Recipient[]): string {
  if (!recipients || recipients.length === 0) return '(none)';
  return recipients.map((r) => r.emailAddress.address).join(', ');
}
