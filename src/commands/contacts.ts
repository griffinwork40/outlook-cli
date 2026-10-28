/**
 * `outlook contacts <subcommand>` — read-only access to the user's contacts
 * via Microsoft Graph v1.0 `/me/contacts`.
 *
 * Subcommand is positional 0 (`list` is the default when absent).
 * Unknown subcommands are a UsageError listing the valid ones.
 *
 * `mutating: false` — entirely read-only. No `guardMutation` needed.
 *
 * Search uses `$filter=startswith(displayName,'<query>')` rather than
 * `$search="displayName:{query}"` — the latter requires a
 * `ConsistencyLevel: eventual` header which would need client.ts changes.
 *
 * NOTE: the `Contacts.Read` scope is required. Users upgrading from a
 * previous version must re-login (`outlook auth logout && outlook auth login`)
 * to pick up the new scope.
 */

import { emit, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import { renderContact, renderContactList } from '../output.js';
import { collectItems } from '../paginate.js';
import type { Contact } from '../types.js';

const SUBCOMMANDS = ['list', 'get', 'search'] as const;

const DEFAULT_SELECT =
  'id,displayName,emailAddresses,mobilePhone,businessPhones,homePhones,companyName,jobTitle';
const DEFAULT_ORDERBY = 'displayName asc';
const MAX_TOP = 1000;

function readTop(ctx: CommandContext): number {
  if (!ctx.args.has('top')) return 25;
  const top = ctx.args.num('top');
  if (top === undefined || !Number.isInteger(top) || top < 1 || top > MAX_TOP) {
    throw new UsageError(`--top must be an integer between 1 and ${MAX_TOP}`);
  }
  return top;
}

function readLimit(ctx: CommandContext, fallback: number): number {
  if (!ctx.args.has('limit')) return fallback;
  const limit = ctx.args.num('limit');
  if (limit === undefined || !Number.isInteger(limit) || limit < 1 || limit > MAX_TOP) {
    throw new UsageError(`--limit must be an integer between 1 and ${MAX_TOP}`);
  }
  return limit;
}

async function listHandler(ctx: CommandContext): Promise<void> {
  const top = readTop(ctx);
  const all = ctx.args.bool('all');
  const limit = readLimit(ctx, top);
  const select = ctx.args.str('select') ?? DEFAULT_SELECT;

  const query: Record<string, unknown> = {
    $select: select,
    $orderby: DEFAULT_ORDERBY,
    $top: top,
  };

  const filter = ctx.args.str('filter');
  if (filter) query.$filter = filter;

  const contacts = await collectItems<Contact>(ctx.client, '/me/contacts', {
    query,
    limit,
    all,
  });

  emit(ctx, contacts, () => renderContactList(contacts));
}

async function getHandler(ctx: CommandContext): Promise<void> {
  const id = ctx.args.requirePositional(1, '<contactId>');
  const path = `/me/contacts/${encodeURIComponent(id)}`;

  const contact = await ctx.client.get<Contact>(path);

  emit(ctx, contact, () => renderContact(contact));
}

async function searchHandler(ctx: CommandContext): Promise<void> {
  const query = ctx.args.requirePositional(1, '<query>');
  const top = readTop(ctx);

  // Use $filter with startswith — no ConsistencyLevel header needed.
  const escaped = query.replace(/'/g, "''");
  const filterExpr = `startswith(displayName,'${escaped}')`;

  const contacts = await collectItems<Contact>(ctx.client, '/me/contacts', {
    query: {
      $select: DEFAULT_SELECT,
      $filter: filterExpr,
      $top: top,
    },
    limit: top,
  });

  emit(ctx, contacts, () => renderContactList(contacts));
}

async function contactsHandler(ctx: CommandContext): Promise<void> {
  const subcommand = ctx.args.positionals[0] ?? 'list';
  switch (subcommand) {
    case 'list':
      return listHandler(ctx);
    case 'get':
      return getHandler(ctx);
    case 'search':
      return searchHandler(ctx);
    default:
      throw new UsageError(
        `Unknown "outlook contacts" subcommand: "${subcommand}". Valid subcommands: ${SUBCOMMANDS.join(', ')}.`,
      );
  }
}

export const contactsCommand: CommandSpec = {
  name: 'contacts',
  summary: 'List, get, or search contacts (read-only)',
  usage:
    'outlook contacts [list|get|search] [<id|query>] [--top <n>] [--all] [--limit <n>] [--select <fields>] [--filter <expr>] [--json]',
  mutating: false,
  handler: contactsHandler,
};
