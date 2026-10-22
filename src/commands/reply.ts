/**
 * `outlook reply` / `reply-all` / `forward`.
 *
 * Two paths, chosen by whether the resolved signature has an HTML form:
 *
 * TEXT-ONLY (or no signature):
 *   POST /me/messages/{id}/reply|replyAll|forward  { comment }
 *   Single request; existing behaviour unchanged.
 *
 * HTML SIGNATURE:
 *   1. POST /me/messages/{id}/createReply|createReplyAll|createForward
 *      → Exchange creates a draft with the quoted body and sets threading headers.
 *   2. GET the draft, insert user text + signature above the Exchange quote,
 *      PATCH /me/messages/{draftId}  { body }
 *   3. POST /me/messages/{draftId}/send
 *
 *   On PATCH or send failure the draft is deleted and the error reported.
 *   If the delete also fails, the draft id is surfaced so the user can clean up.
 *   --dry-run shows all three requests without touching the network.
 *
 * Why the split rather than always using the 3-step path?
 *   The one-shot action is one request, atomic, and battle-tested.  For users
 *   with only a plain-text signature there is no reason to trade that in for a
 *   non-atomic three-request flow.  The split keeps the fast path fast and
 *   limits the new surface to where it is actually needed.
 *
 * docs/02 §2: replyAll documents only `comment`; the surrounding prose implies
 * a `message` object may also be accepted but is unverified.  The three-step
 * path uses createReplyAll (a separate action) which is documented to return a
 * draft, so we avoid the ambiguity entirely.
 */

import { emit, guardMutation, type CommandContext, type CommandSpec } from '../command.js';
import { renderRequest } from '../dryRun.js';
import { UsageError } from '../errors.js';
import { appendTextSignature, resolveSignature } from '../signature.js';
import { describeRecipients, recipientsFromArgs } from './compose.js';
import {
  buildReplyDraftSpecs,
  executeReplyDraft,
  type CreateAction,
} from './replyDraft.js';

type ReplyKind = 'reply' | 'replyAll' | 'forward';

const ACTION_PATH: Record<ReplyKind, string> = {
  reply: 'reply',
  replyAll: 'replyAll',
  forward: 'forward',
};

const CREATE_ACTION: Record<ReplyKind, CreateAction> = {
  reply: 'createReply',
  replyAll: 'createReplyAll',
  forward: 'createForward',
};

async function runReplyAction(
  ctx: CommandContext,
  kind: ReplyKind,
  label: string,
): Promise<void> {
  const messageId = ctx.args.requirePositional(0, '<messageId>');
  const userText = ctx.args.str('body') ?? '';
  const signature = await resolveSignature(ctx.args, ctx.env);

  // Forward always needs recipients regardless of path.
  const toRecipients = kind === 'forward' ? recipientsFromArgs(ctx.args, 'to') : undefined;
  if (kind === 'forward' && (!toRecipients || toRecipients.length === 0)) {
    throw new UsageError(
      'forward requires at least one recipient. Use --to <address>.',
    );
  }

  // Choose path: HTML signature → 3-step draft flow; otherwise one-shot comment.
  const useHtmlFlow = signature !== undefined && signature.html !== undefined;

  if (useHtmlFlow) {
    await runHtmlFlow(ctx, kind, label, messageId, userText, signature!, toRecipients);
    return;
  }

  // -------------------------------------------------------------------------
  // One-shot comment path (text-only or no signature).
  // -------------------------------------------------------------------------
  const comment =
    signature === undefined ? userText : appendTextSignature(userText, signature);

  const payload: Record<string, unknown> = { comment };
  if (toRecipients) payload['toRecipients'] = toRecipients;

  const path = `/me/messages/${encodeURIComponent(messageId)}/${ACTION_PATH[kind]}`;
  const spec = ctx.client.buildSpec('POST', path, payload);

  const done = await guardMutation(ctx, spec, async () => {
    await ctx.client.post<void>(path, payload);
    return true;
  });

  if (done === undefined) return; // --dry-run

  emit(
    ctx,
    {
      action: kind,
      messageId,
      sent: true,
      signatureApplied: signature !== undefined,
      htmlSignature: false,
    },
    () =>
      `${label} sent for message ${messageId}` +
      (toRecipients ? ` to ${describeRecipients(toRecipients)}` : ''),
  );
}

/**
 * Three-step HTML-signature flow.
 * --dry-run renders all three requests and returns without touching the network.
 */
async function runHtmlFlow(
  ctx: CommandContext,
  kind: ReplyKind,
  label: string,
  messageId: string,
  userText: string,
  signature: NonNullable<Awaited<ReturnType<typeof resolveSignature>>>,
  toRecipients: ReturnType<typeof recipientsFromArgs> | undefined,
): Promise<void> {
  const specs = buildReplyDraftSpecs(ctx.client, {
    messageId,
    action: CREATE_ACTION[kind],
    userText,
    signature,
    toRecipients,
  });

  if (ctx.dryRun) {
    ctx.out('DRY RUN — nothing was sent. The three requests below are what would go out:');
    ctx.out('');
    ctx.out('--- Step 1: create draft ---');
    ctx.out(renderRequest(specs.createSpec));
    ctx.out('');
    ctx.out('--- Step 2: patch body (representative; actual body depends on Exchange draft) ---');
    ctx.out(renderRequest(specs.patchSpec));
    ctx.out('');
    ctx.out('--- Step 3: send draft ---');
    ctx.out(renderRequest(specs.sendSpec));
    return;
  }

  // Guard: check write scope via the create spec (same check guardMutation does).
  // We do it once here rather than once per step.
  const dummy = await guardMutation(ctx, specs.createSpec, async () => true);
  if (dummy === undefined) return; // unreachable when dryRun=false, but satisfies TS

  const draftId = await executeReplyDraft(ctx.client, {
    messageId,
    action: CREATE_ACTION[kind],
    userText,
    signature,
    toRecipients,
  });

  emit(
    ctx,
    {
      action: kind,
      messageId,
      sent: true,
      signatureApplied: true,
      htmlSignature: true,
      draftId,
    },
    () =>
      `${label} sent for message ${messageId}` +
      (toRecipients ? ` to ${describeRecipients(toRecipients)}` : ''),
  );
}

export const replyCommand: CommandSpec = {
  name: 'reply',
  summary: 'Reply to the sender of a message',
  usage:
    'outlook reply <messageId> [--body <text>] ' +
    '[--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'reply', 'Reply'),
};

export const replyAllCommand: CommandSpec = {
  name: 'reply-all',
  summary: 'Reply to all recipients of a message',
  usage:
    'outlook reply-all <messageId> [--body <text>] ' +
    '[--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'replyAll', 'Reply-all'),
};

export const forwardCommand: CommandSpec = {
  name: 'forward',
  summary: 'Forward a message to other recipients',
  usage:
    'outlook forward <messageId> --to <addr> [--to <addr>...] ' +
    '[--body <text>] [--signature <path> | --no-signature] [--dry-run] [--json]',
  mutating: true,
  handler: (ctx) => runReplyAction(ctx, 'forward', 'Forward'),
};
