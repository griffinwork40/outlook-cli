/**
 * Enforcement of the OUTLOOK_POLICY mutation policy (policy.ts).
 *
 * Called only from guardMutation (command.ts), the chokepoint every
 * state-changing Graph request already passes through, so no command can skip
 * it. With OUTLOOK_POLICY unset every function here is a no-op and behavior is
 * exactly as before.
 *
 * This is a guard INSIDE the CLI, not an operating-system boundary: a process
 * that can run arbitrary code as the same user can read the token cache and
 * call Graph directly. Its job is to keep an agent that drives the CLI inside
 * the lines it was given, and to fail closed when the policy cannot be read.
 */

import { dirname, resolve } from 'node:path';
import type { CommandContext } from './command.js';
import { PolicyRefusedError } from './errors.js';
import { POLICY_ENV, loadDenyList, loadPolicy, sameFolder, type MutationPolicy } from './policy.js';
import type { RequestSpec } from './types.js';

export interface PolicyVerdict {
  policyPath: string;
  allowed: boolean;
  /** Which clause refused: policy-file, allowCommands, request-path, move.allowedDestinations, move.denySenders. */
  rule?: string;
  reason?: string;
  /** Deny-list status for a move into a protected destination. */
  senderCheck?: 'passed' | 'skipped-dry-run';
}

export interface PolicyInput {
  env: Readonly<Record<string, string | undefined>> | undefined;
  /** Registered command name, e.g. "move" or "draft". */
  command: string | undefined;
  /** positionals[0] — the subcommand for draft/folders/rules. */
  subcommand: string | undefined;
  spec: RequestSpec;
  dryRun: boolean;
  /** GET the message's sender address. Never called under dry-run. */
  lookupSender: (messageId: string) => Promise<string | undefined>;
  now?: () => number;
}

export function policyPath(
  env: Readonly<Record<string, string | undefined>> | undefined,
): string | undefined {
  const value = env?.[POLICY_ENV]?.trim();
  return value ? value : undefined;
}

function normalizeCommand(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/** "draft create" is allowed by an entry of either "draft create" or "draft". */
function commandAllowed(policy: MutationPolicy, command: string, sub: string | undefined): boolean {
  const allowed = new Set(policy.allowCommands.map(normalizeCommand));
  const base = normalizeCommand(command);
  if (allowed.has(base)) return true;
  return sub !== undefined && allowed.has(normalizeCommand(`${base} ${sub}`));
}

/** Path part of the request URL, before any query string, without normalization. */
function rawPath(url: string): string {
  return url.replace(/^https?:\/\/[^/]*/i, '').split(/[?#]/)[0] ?? '';
}

/**
 * Message ids go into request paths unescaped, so an id such as
 * `../mailFolders/inbox` would turn "move one message" into "move a whole
 * folder" once the URL is normalized. Under a policy, refuse any path the
 * URL parser would rewrite.
 */
function unsafePathReason(url: string): string | undefined {
  for (const segment of rawPath(url).split('/')) {
    let decoded: string;
    try {
      decoded = decodeURIComponent(segment);
    } catch {
      return 'request path contains a malformed percent-encoding';
    }
    if (decoded === '.' || decoded === '..' || decoded.includes('\\') || decoded.includes('/')) {
      return 'request path contains a dot segment, backslash or encoded slash';
    }
  }
  return undefined;
}

interface MoveTarget {
  destination: string;
  messageId: string | undefined;
}

/** A POST ending in /move (message or folder move); undefined for anything else. */
function moveTarget(spec: RequestSpec): MoveTarget | undefined {
  if (spec.method.toUpperCase() !== 'POST') return undefined;
  const path = rawPath(spec.url);
  if (!path.endsWith('/move')) return undefined;
  const body = spec.body as { destinationId?: unknown } | undefined;
  const destination = typeof body?.destinationId === 'string' ? body.destinationId : '';
  const match = /\/me\/messages\/([^/]+)\/move$/.exec(path);
  return { destination, messageId: match?.[1] };
}

/** Undefined when no policy is configured; otherwise the verdict for this request. */
export async function evaluatePolicy(input: PolicyInput): Promise<PolicyVerdict | undefined> {
  const path = policyPath(input.env);
  if (path === undefined) return undefined;
  const refuse = (rule: string, reason: string): PolicyVerdict => ({
    policyPath: path,
    allowed: false,
    rule,
    reason,
  });

  const loaded = await loadPolicy(path);
  if (!loaded.ok) return refuse('policy-file', loaded.reason);
  const policy = loaded.value;

  if (input.command === undefined) {
    return refuse('allowCommands', 'cannot determine which command issued this request');
  }
  if (!commandAllowed(policy, input.command, input.subcommand)) {
    const label = input.subcommand ? `${input.command} ${input.subcommand}` : input.command;
    return refuse('allowCommands', `'${label}' is not in allowCommands`);
  }

  const pathProblem = unsafePathReason(input.spec.url);
  if (pathProblem !== undefined) return refuse('request-path', pathProblem);

  const move = moveTarget(input.spec);
  if (move === undefined) return { policyPath: path, allowed: true };
  const { destination, messageId } = move;
  if (destination === '') return refuse('move.allowedDestinations', 'move request has no destinationId');

  const allowedDestinations = policy.move?.allowedDestinations;
  if (allowedDestinations !== undefined && !allowedDestinations.some((d) => sameFolder(d, destination))) {
    return refuse(
      'move.allowedDestinations',
      `destination '${destination}' is not in move.allowedDestinations`,
    );
  }

  const deny = policy.move?.denySenders;
  if (deny === undefined || !deny.destinations.some((d) => sameFolder(d, destination))) {
    return { policyPath: path, allowed: true };
  }

  const now = input.now ?? Date.now;
  const listFile = resolve(dirname(path), deny.file);
  const list = await loadDenyList(listFile, deny.maxAgeHours, now());
  if (!list.ok) return refuse('move.denySenders', `${list.reason}; moves to '${destination}' are refused`);
  if (messageId === undefined) {
    return refuse('move.denySenders', 'cannot identify the message, so its sender cannot be checked');
  }
  if (input.dryRun) return { policyPath: path, allowed: true, senderCheck: 'skipped-dry-run' };

  // Errors (auth, 404, throttling) propagate: the move does not happen.
  const sender = (await input.lookupSender(messageId))?.trim().toLowerCase();
  if (!sender) return refuse('move.denySenders', 'the message has no sender address to check');
  if (list.value.has(sender)) {
    return refuse('move.denySenders', `sender ${sender} is on the deny list for '${destination}'`);
  }
  return { policyPath: path, allowed: true, senderCheck: 'passed' };
}

async function verdictFor(ctx: CommandContext, spec: RequestSpec): Promise<PolicyVerdict | undefined> {
  return evaluatePolicy({
    env: ctx.env,
    command: ctx.command,
    subcommand: ctx.args.positionals[0],
    spec,
    dryRun: ctx.dryRun,
    lookupSender: async (messageId) => {
      const message = await ctx.client.get<{ from?: { emailAddress?: { address?: string } } }>(
        `/me/messages/${messageId}`,
        { $select: 'from' },
      );
      return message?.from?.emailAddress?.address;
    },
  });
}

/** Throws PolicyRefusedError (exit 4) when a configured policy refuses the request. */
export async function enforcePolicy(ctx: CommandContext, spec: RequestSpec): Promise<void> {
  const verdict = await verdictFor(ctx, spec);
  if (verdict === undefined || verdict.allowed) return;
  const rule = verdict.rule ?? 'policy';
  const reason = verdict.reason ?? 'refused';
  if (ctx.json) {
    ctx.out(
      JSON.stringify(
        { error: { code: 'PolicyRefused', rule, message: reason, policy: verdict.policyPath } },
        null,
        2,
      ),
    );
  }
  throw new PolicyRefusedError(rule, reason);
}

/** For --dry-run: the verdict without any network call (undefined when no policy is set). */
export async function dryRunVerdict(
  ctx: CommandContext,
  spec: RequestSpec,
): Promise<PolicyVerdict | undefined> {
  return verdictFor({ ...ctx, dryRun: true }, spec);
}

export function renderVerdict(verdict: PolicyVerdict): string {
  const head = `Policy (${verdict.policyPath}):`;
  if (!verdict.allowed) return `${head} would REFUSE (${verdict.rule}): ${verdict.reason}`;
  if (verdict.senderCheck === 'skipped-dry-run') {
    return (
      `${head} would ALLOW unless the message's sender is on the deny list ` +
      `(the sender lookup is a network call, so --dry-run skipped it).`
    );
  }
  return `${head} would ALLOW this request.`;
}
