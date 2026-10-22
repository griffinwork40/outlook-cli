/**
 * `outlook rules` — list and create inbox message rules.
 *
 * Graph endpoint (list): `GET /me/mailFolders/inbox/messageRules`
 * Graph endpoint (create): `POST /me/mailFolders/inbox/messageRules`
 * (see https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules)
 * (see https://learn.microsoft.com/en-us/graph/api/mailfolder-post-messagerules)
 *
 * Each messageRule has: displayName, sequence, isEnabled, isReadOnly,
 * conditions, actions, exceptions. `list` is read-only — no guardMutation
 * needed, no --dry-run. `create` is mutating: it creates a new server-side
 * inbox rule via POST and goes through `guardMutation` for --dry-run safety.
 *
 * Subcommand dispatch on positional 0 (`list` is the default):
 *   outlook rules [list] [--json]
 *   outlook rules create --name <text> --folder <id>
 *       [--from <addr>...] [--sender-contains <domain>...]
 *       [--mark-read] [--stop-processing] [--enabled]
 *       [--dry-run] [--json]
 *
 * SAFETY: `create` is UNSAFE to retry — Graph applies no uniqueness constraint
 * on rule displayName; a blindly-retried POST creates a duplicate. Wrapped in
 * `guardMutation` and never auto-retried (POST is not in RETRYABLE_METHODS).
 *
 * Scope requirement: `MailboxSettings.ReadWrite` must be consented at login.
 * `MailboxSettings.Read` is sufficient for `list` but not for `create`.
 */

import { emit, guardMutation, type CommandContext, type CommandSpec } from '../command.js';
import { UsageError } from '../errors.js';
import type { GraphCollection } from '../types.js';

/* ---------- Graph response types (subset we render) ---------- */

interface RuleEmailAddress {
  emailAddress?: { address?: string; name?: string };
}

/** Conditions or exceptions on a message rule. */
export interface RuleConditions {
  fromAddresses?: RuleEmailAddress[];
  sentToAddresses?: RuleEmailAddress[];
  subjectContains?: string[];
  bodyContains?: string[];
  headerContains?: string[];
  senderContains?: string[];
  recipientContains?: string[];
  bodyOrSubjectContains?: string[];
  hasAttachments?: boolean;
  importance?: string;
  isAutomatic?: boolean;
  [key: string]: unknown;
}

/** Actions performed when a rule matches. */
export interface RuleActions {
  moveToFolder?: string;
  copyToFolder?: string;
  delete?: boolean;
  permanentDelete?: boolean;
  markAsRead?: boolean;
  markImportance?: string;
  forwardTo?: RuleEmailAddress[];
  forwardAsAttachmentTo?: RuleEmailAddress[];
  redirectTo?: RuleEmailAddress[];
  stopProcessingRules?: boolean;
  [key: string]: unknown;
}

export interface MessageRule {
  id?: string;
  displayName?: string;
  sequence?: number;
  isEnabled?: boolean;
  isReadOnly?: boolean;
  conditions?: RuleConditions;
  actions?: RuleActions;
  exceptions?: RuleConditions;
}

/* ---------- Human-readable summarisers ---------- */

function addrList(addrs?: RuleEmailAddress[]): string {
  if (!addrs || addrs.length === 0) return '';
  return addrs
    .map((a) => a.emailAddress?.address ?? a.emailAddress?.name ?? '?')
    .join(', ');
}

/** Short human string describing which messages match. */
export function summariseConditions(c?: RuleConditions): string {
  if (!c) return '(any)';
  const parts: string[] = [];

  const from = addrList(c.fromAddresses);
  if (from) parts.push(`from: ${from}`);

  const to = addrList(c.sentToAddresses);
  if (to) parts.push(`to: ${to}`);

  if (c.subjectContains?.length) parts.push(`subject contains: ${c.subjectContains.join(', ')}`);
  if (c.bodyContains?.length) parts.push(`body contains: ${c.bodyContains.join(', ')}`);
  if (c.senderContains?.length) parts.push(`sender contains: ${c.senderContains.join(', ')}`);
  if (c.hasAttachments) parts.push('has attachments');
  if (c.importance) parts.push(`importance: ${c.importance}`);

  return parts.length > 0 ? parts.join('; ') : '(any)';
}

/** Short human string describing what the rule does. */
export function summariseActions(a?: RuleActions): string {
  if (!a) return '(none)';
  const parts: string[] = [];

  if (a.moveToFolder) parts.push(`move to [${a.moveToFolder}]`);
  if (a.copyToFolder) parts.push(`copy to [${a.copyToFolder}]`);
  if (a.delete) parts.push('delete');
  if (a.permanentDelete) parts.push('permanent delete');
  if (a.markAsRead) parts.push('mark read');
  if (a.markImportance) parts.push(`mark ${a.markImportance}`);

  const fwd = addrList(a.forwardTo);
  if (fwd) parts.push(`forward to: ${fwd}`);
  const redir = addrList(a.redirectTo);
  if (redir) parts.push(`redirect to: ${redir}`);

  if (a.stopProcessingRules) parts.push('stop processing');

  return parts.length > 0 ? parts.join('; ') : '(none)';
}

export function renderRules(rules: MessageRule[]): string {
  if (rules.length === 0) return '(no rules)';

  const header = [
    '#'.padEnd(3),
    'ENABLED'.padEnd(8),
    'NAME'.padEnd(28),
    'CONDITIONS'.padEnd(36),
    'ACTIONS',
  ].join(' ');

  const rows = rules.map((r) => [
    String(r.sequence ?? '-').padEnd(3),
    (r.isEnabled ? 'yes' : 'no').padEnd(8),
    (r.displayName ?? '(unnamed)').slice(0, 28).padEnd(28),
    summariseConditions(r.conditions).slice(0, 36).padEnd(36),
    summariseActions(r.actions),
  ].join(' '));

  return [header, ...rows].join('\n');
}

/* ---------- Create handler ---------- */

/**
 * Build the Graph messageRule request body from CLI flags.
 *
 * Conditions:
 *   --from <addr>         → conditions.fromAddresses (repeatable)
 *   --sender-contains <s> → conditions.senderContains (repeatable, for domain matching)
 *
 * Actions:
 *   --folder <id>         → actions.moveToFolder (required; opaque folder id)
 *   --mark-read           → actions.markAsRead = true
 *   --stop-processing     → actions.stopProcessingRules = true (default: true)
 *
 * Meta:
 *   --name <text>         → displayName (required)
 *   --enabled             → isEnabled (default: true)
 *   --sequence <n>        → sequence (optional; Graph auto-assigns if omitted)
 */
function buildRuleBody(ctx: CommandContext): Record<string, unknown> {
  const name = ctx.args.require('name');
  const folderId = ctx.args.require('folder');

  // Conditions — at least one of --from or --sender-contains is required
  const fromAddrs = ctx.args.strAll('from');
  const senderContains = ctx.args.strAll('sender-contains');

  if (fromAddrs.length === 0 && senderContains.length === 0) {
    throw new UsageError(
      'At least one condition is required: --from <address> and/or --sender-contains <domain>. ' +
        'A rule with no condition would match every message.',
    );
  }

  const conditions: RuleConditions = {};
  if (fromAddrs.length > 0) {
    conditions.fromAddresses = fromAddrs.map((addr) => ({
      emailAddress: { address: addr },
    }));
  }
  if (senderContains.length > 0) {
    conditions.senderContains = senderContains;
  }

  // Actions
  const actions: RuleActions = {
    moveToFolder: folderId,
  };
  if (ctx.args.bool('mark-read')) {
    actions.markAsRead = true;
  }
  // Default to stopProcessingRules: true unless --no-stop-processing
  if (!ctx.args.has('stop-processing') || ctx.args.bool('stop-processing')) {
    actions.stopProcessingRules = true;
  }

  // Meta
  const body: Record<string, unknown> = {
    displayName: name,
    conditions,
    actions,
    isEnabled: !ctx.args.has('enabled') || ctx.args.bool('enabled'),
  };

  const seq = ctx.args.num('sequence');
  if (seq !== undefined) body.sequence = seq;

  return body;
}

async function createHandler(ctx: CommandContext): Promise<void> {
  const body = buildRuleBody(ctx);
  const spec = ctx.client.buildSpec('POST', RULES_PATH, body);

  const created = await guardMutation(ctx, spec, async () =>
    ctx.client.post<MessageRule>(RULES_PATH, body),
  );

  if (created === undefined) return; // --dry-run already reported

  emit(ctx, created, () => {
    const name = created.displayName ?? '(unnamed)';
    const seq = created.sequence ?? '?';
    const cond = summariseConditions(created.conditions);
    const acts = summariseActions(created.actions);
    return `Created rule #${seq} "${name}"\n  Conditions: ${cond}\n  Actions: ${acts}`;
  });
}

/* ---------- Command handler ---------- */

const RULES_PATH = '/me/mailFolders/inbox/messageRules';

async function rulesHandler(ctx: CommandContext): Promise<void> {
  const sub = ctx.args.positionals[0] ?? 'list';

  switch (sub) {
    case 'list': {
      const response = await ctx.client.get<GraphCollection<MessageRule>>(RULES_PATH);
      const rules = response.value;
      emit(ctx, rules, () => renderRules(rules));
      return;
    }
    case 'create':
      return createHandler(ctx);
    default: {
      throw new UsageError(
        `Unknown rules subcommand "${sub}". Expected "list" or "create".`,
      );
    }
  }
}

export const rulesCommand: CommandSpec = {
  name: 'rules',
  summary: 'List or create inbox message rules',
  usage:
    'outlook rules [list] [--json]\n' +
    '       outlook rules create --name <text> --folder <id>\n' +
    '           [--from <addr>...] [--sender-contains <domain>...]\n' +
    '           [--mark-read] [--stop-processing] [--enabled]\n' +
    '           (--stop-processing defaults to on; use --no-stop-processing to disable)\n' +
    '           [--dry-run] [--json]',
  mutating: true,
  handler: rulesHandler,
};
