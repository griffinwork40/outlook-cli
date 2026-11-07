/**
 * outlook delete / move / purge.
 *
 * SAFETY (docs/02-write-endpoints.md §4): Microsoft's own docs confirm that
 * `POST /me/messages/{id}/move` -> {"destinationId":"deleteditems"} lands the
 * message in the visible, user-browsable Deleted Items folder — recoverable
 * in the ordinary Outlook sense. The docs are silent on whether a raw
 * `DELETE /me/messages/{id}` is recoverable at all. An undocumented
 * irreversible-or-not operation is treated as irreversible by default here.
 *
 * `delete` therefore performs the MOVE, never the raw DELETE. Only `purge`
 * issues the raw DELETE, and only when the caller passes --yes.
 *
 * IMPORTANT: per §4's own worked example, a successful move returns a NEW
 * message id — the id used in the request no longer resolves afterward.
 * Callers must treat the response's id as the message's identity going
 * forward, not the id they passed in.
 */

import type { CommandContext, CommandSpec } from '../command.js';
import { guardMutation, emit } from '../command.js';
import { isWellKnownFolder } from '../types.js';
import type { Message } from '../types.js';
import { UsageError } from '../errors.js';

const DELETED_ITEMS = 'deleteditems';

function messagePath(id: string): string {
  return `/me/messages/${encodeURIComponent(id)}`;
}

/** Well-known names are documented as lowercase literals; a raw folder id is opaque and case-sensitive. */
function resolveDestinationId(folder: string): string {
  return isWellKnownFolder(folder) ? folder.toLowerCase() : folder;
}

async function moveMessage(ctx: CommandContext, id: string, destinationId: string): Promise<void> {
  const path = `${messagePath(id)}/move`;
  const body = { destinationId };
  const spec = ctx.client.buildSpec('POST', path, body);

  await guardMutation(ctx, spec, async () => {
    const moved = await ctx.client.post<Message>(path, body);
    emit(
      ctx,
      { movedFrom: id, destinationId, message: moved },
      () =>
        `Moved message ${id} -> ${destinationId}. New id: ${moved.id} ` +
        `(the old id no longer resolves — use the new id for anything further).`,
    );
  });
}

async function deleteHandler(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(0, '<id>');
  await moveMessage(ctx, id, DELETED_ITEMS);
}

export const deleteCommand: CommandSpec = {
  name: 'delete',
  summary: 'Move a message to Deleted Items (recoverable). The safe default "delete".',
  usage: 'outlook delete <id> [--dry-run] [--json]',
  mutating: true,
  handler: deleteHandler,
};

async function moveHandler(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(0, '<id>');
  const folder = ctx.args.require('folder');
  await moveMessage(ctx, id, resolveDestinationId(folder));
}

export const moveCommand: CommandSpec = {
  name: 'move',
  summary: 'Move a message to another folder (well-known name or folder id).',
  usage: 'outlook move <id> --folder <name|id> [--dry-run] [--json]',
  mutating: true,
  handler: moveHandler,
};

async function purgeHandler(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(0, '<id>');

  if (!ctx.args.bool('yes')) {
    throw new UsageError(
      `outlook purge issues a raw DELETE /me/messages/${id}. Microsoft's docs ` +
        `never state whether this is recoverable — it is treated as permanent. ` +
        `Pass --yes to confirm you understand this, or run \`outlook delete ${id}\` ` +
        `instead, which moves the message to Deleted Items and IS recoverable.`,
    );
  }

  const path = messagePath(id);
  const spec = ctx.client.buildSpec('DELETE', path);

  await guardMutation(ctx, spec, async () => {
    await ctx.client.del(path);
    emit(
      ctx,
      { purged: true, id },
      () =>
        `Purged message ${id} via raw DELETE. Recoverability is not documented ` +
        `by Microsoft — treat this as permanent.`,
    );
  });
}

export const purgeCommand: CommandSpec = {
  name: 'purge',
  summary: 'Permanently delete a message via raw DELETE. Requires --yes.',
  usage: 'outlook purge <id> --yes [--dry-run] [--json]',
  mutating: true,
  handler: purgeHandler,
};
