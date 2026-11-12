/**
 * `outlook draft list|create|update|send`.
 *
 * The draft lifecycle is the SAFE way to send mail: create, inspect, then
 * send as a separate deliberate step. `draft send` is irreversible and is
 * guarded; `draft list` is a plain read.
 */

import { emit, guardMutation, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { collectItems } from '../paginate.js';
import { renderMessageList } from '../output.js';
import { resolveSignature, signatureFlagRequested } from '../signature.js';
import type { Message } from '../types.js';
import { composeMessage, signMessage } from './compose.js';

const DRAFTS_PATH = '/me/mailFolders/drafts/messages';
const SUBCOMMANDS = ['list', 'create', 'update', 'send'] as const;

async function listDrafts(ctx: CommandContext): Promise<void> {
  const limit = ctx.args.num('limit');
  const messages = await collectItems<Message>(ctx.client, DRAFTS_PATH, {
    query: {
      $select: 'id,subject,toRecipients,lastModifiedDateTime,isDraft',
      $orderby: 'lastModifiedDateTime desc',
    },
    ...(limit === undefined ? {} : { limit }),
    ...(ctx.args.bool('all') ? { all: true } : {}),
  });

  emit(ctx, messages, () => renderMessageList(messages));
}

async function createDraft(ctx: CommandContext): Promise<void> {
  // Compose and check 'nothing to create' before reading the
  // signature file, so an empty draft surfaces the right error.
  const message = await composeMessage(ctx.args);

  if (Object.keys(message).length === 0) {
    throw new UsageError(
      'Nothing to create. Provide at least --subject, --body or --to.',
    );
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

  const spec = ctx.client.buildSpec('POST', '/me/messages', message);

  const created = await guardMutation(ctx, spec, () =>
    ctx.client.post<Message>('/me/messages', message),
  );

  if (created === undefined) return;

  emit(
    ctx,
    {
      ...created,
      signatureApplied: signature !== undefined,
      inlineSignatureImages: signature?.inlineAttachments?.length ?? 0,
    },
    () => `Draft created: ${created.id}\nSend it with:  outlook draft send ${created.id}`,
  );
}

async function updateDraft(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(1, '<draftId>');
  // Never sign on update: the draft was signed at create time, and a replaced
  // body would otherwise gain a second copy on every edit.
  if (signatureFlagRequested(ctx.args)) {
    throw new UsageError(
      'draft update never adds a signature (it would duplicate the one added ' +
        'at create time). Include it in --body yourself if you need it.',
    );
  }
  const message = await composeMessage(ctx.args);

  if (Object.keys(message).length === 0) {
    throw new UsageError(
      'Nothing to update. Provide at least one of --subject, --body, --to, --cc, --bcc.',
    );
  }

  const path = `/me/messages/${encodeURIComponent(id)}`;
  const spec = ctx.client.buildSpec('PATCH', path, message);

  const updated = await guardMutation(ctx, spec, () =>
    ctx.client.patch<Message>(path, message),
  );

  if (updated === undefined) return;
  emit(ctx, updated, () => `Draft ${id} updated.`);
}

async function sendDraft(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(1, '<draftId>');
  const path = `/me/messages/${encodeURIComponent(id)}/send`;
  const spec = ctx.client.buildSpec('POST', path, undefined);

  const sent = await guardMutation(ctx, spec, async () => {
    await ctx.client.post<void>(path);
    return true;
  });

  if (sent === undefined) return;
  emit(ctx, { sent: true, draftId: id }, () => `Draft ${id} sent.`);
}

export const draftCommand: CommandSpec = {
  name: 'draft',
  summary: 'Create, update, list and send drafts',
  usage:
    'outlook draft list [--limit N] [--all] [--json]\n' +
    '  outlook draft create --subject <s> [--to <addr>...] [--body <text>] [--html] [--attach <path>...] [--dry-run]\n' +
    '  outlook draft update <draftId> [--subject <s>] [--body <text>] [--to <addr>...] [--dry-run]\n' +
    '  outlook draft send <draftId> [--dry-run]',
  mutating: true,

  async handler(ctx) {
    const sub = ctx.args.positionals[0];

    switch (sub) {
      case 'list':
        return listDrafts(ctx);
      case 'create':
        return createDraft(ctx);
      case 'update':
        return updateDraft(ctx);
      case 'send':
        return sendDraft(ctx);
      default:
        throw new UsageError(
          `Unknown draft subcommand${sub ? ` "${sub}"` : ''}. ` +
            `Valid subcommands: ${SUBCOMMANDS.join(', ')}.`,
        );
    }
  },
};
