/**
 * `outlook send` — POST /me/sendMail.
 *
 * This is the most irreversible command in the CLI: it delivers real mail to
 * real people and cannot be recalled. Two invariants:
 *   1. The Graph call is wrapped in `guardMutation`, so --dry-run short-circuits
 *      before any network I/O.
 *   2. No retry. client.ts refuses to retry POST because sendMail carries no
 *      idempotency guarantee (docs/02 §8) — a retry double-sends.
 *
 * sendMail returns 202 Accepted with NO body, so there is no Message to parse.
 */

import { emit, guardMutation, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { resolveSignature } from '../signature.js';
import {
  composeMessage,
  describeRecipients,
  saveToSentItems,
  signMessage,
} from './compose.js';

export const sendCommand: CommandSpec = {
  name: 'send',
  summary: 'Send an email',
  usage:
    'outlook send --to <addr> [--to <addr>...] --subject <text> ' +
    '[--cc <addr>] [--bcc <addr>] [--body <text> | --body-file <path>] ' +
    '[--html] [--attach <path>...] [--signature <path> | --no-signature] ' +
    '[--no-save-sent] [--dry-run] [--json]',
  mutating: true,

  async handler(ctx) {
    // Compose first, validate --to/--subject, then resolve + apply
    // signature — so a missing recipient surfaces before a bad signature path.
    const message = await composeMessage(ctx.args);

    if (!message.toRecipients || message.toRecipients.length === 0) {
      throw new UsageError(
        'At least one recipient is required. Use --to <address> ' +
          '(repeatable, or comma-separated).',
      );
    }
    if (message.subject === undefined || message.subject.trim() === '') {
      throw new UsageError('--subject is required.');
    }

    const signature = await resolveSignature(ctx.args, ctx.env);
    signMessage(message, signature, ctx.args);

    // Merge any inline signature images into the message attachment list.
    if (signature?.inlineAttachments?.length) {
      message.attachments = [
        ...(message.attachments ?? []),
        ...signature.inlineAttachments,
      ];
    }

    const payload = {
      message,
      saveToSentItems: saveToSentItems(ctx.args),
    };

    const spec = ctx.client.buildSpec('POST', '/me/sendMail', payload);

    const sent = await guardMutation(ctx, spec, async () => {
      await ctx.client.post<void>('/me/sendMail', payload);
      return true;
    });

    if (sent === undefined) return; // --dry-run already reported

    const to = describeRecipients(message.toRecipients);
    emit(
      ctx,
      {
        sent: true,
        subject: message.subject,
        to: message.toRecipients.map((r) => r.emailAddress.address),
        attachments: message.attachments?.length ?? 0,
        savedToSentItems: payload.saveToSentItems,
        signatureApplied: signature !== undefined,
        inlineSignatureImages: signature?.inlineAttachments?.length ?? 0,
      },
      () => `Sent "${message.subject}" to ${to}`,
    );
  },
};
