/**
 * The contract every command binds to.
 *
 * The --dry-run short-circuit lives HERE, not in individual commands, so that
 * no command can forget it. Sending mail is irreversible and reaches real
 * people; `guardMutation` is the single chokepoint that makes the safe path
 * the default path.
 */

import type { ParsedArgs } from './args.js';
import type { Authenticator } from './auth.js';
import type { GraphClient } from './client.js';
import { renderRequest } from './dryRun.js';
import { InsufficientScopeError } from './errors.js';
import { dryRunVerdict, enforcePolicy, renderVerdict, type PolicyVerdict } from './policyGuard.js';
import { redactObject } from './redact.js';
import type { RequestSpec } from './types.js';

export interface CommandContext {
  args: ParsedArgs;
  /** --json: emit machine-readable output. */
  json: boolean;
  /** --dry-run: render the request, send nothing. */
  dryRun: boolean;
  /**
   * Registered command name (e.g. "move", "draft"), set by cli.ts. The
   * OUTLOOK_POLICY guard matches it against `allowCommands`; when a policy is
   * active and this is absent, every mutation is refused.
   */
  command?: string;
  client: GraphClient;
  auth: Authenticator;
  out: (line: string) => void;
  err: (line: string) => void;
  /**
   * Environment for per-invocation settings (e.g. signature files). Injected
   * rather than read from process.env so tests are hermetic; cli.ts passes
   * process.env. Absent means "nothing configured".
   */
  env?: Readonly<Record<string, string | undefined>>;
}

export type CommandHandler = (ctx: CommandContext) => Promise<void>;

export interface CommandSpec {
  name: string;
  summary: string;
  usage: string;
  /** True when the command changes remote state — used to document --dry-run. */
  mutating: boolean;
  handler: CommandHandler;
}

/**
 * Wrap every state-changing Graph call. When --dry-run is set, prints the
 * exact request that WOULD have been issued and returns undefined without
 * touching the network. When OUTLOOK_POLICY is set, the dry run also reports
 * whether the policy would allow the request, and a real run is refused
 * (PolicyRefusedError, exit 4) unless the policy allows it.
 */
export async function guardMutation<T>(
  ctx: CommandContext,
  spec: RequestSpec,
  exec: () => Promise<T>,
): Promise<T | undefined> {
  if (ctx.dryRun) {
    const verdict = await dryRunVerdict(ctx, spec);
    if (ctx.json) {
      ctx.out(
        JSON.stringify(
          {
            dryRun: true,
            request: redactSpecForJson(spec),
            ...(verdict === undefined ? {} : { policy: verdictJson(verdict) }),
          },
          null,
          2,
        ),
      );
    } else {
      ctx.out('DRY RUN — nothing was sent. The request below is what would go out:');
      ctx.out('');
      ctx.out(renderRequest(spec));
      if (verdict !== undefined) {
        ctx.out('');
        ctx.out(renderVerdict(verdict));
      }
    }
    return undefined;
  }
  await assertWriteScope(ctx, spec);
  await enforcePolicy(ctx, spec);
  return exec();
}

function verdictJson(verdict: PolicyVerdict): Record<string, unknown> {
  return {
    path: verdict.policyPath,
    allowed: verdict.allowed,
    ...(verdict.rule === undefined ? {} : { rule: verdict.rule }),
    ...(verdict.reason === undefined ? {} : { reason: verdict.reason }),
    ...(verdict.senderCheck === undefined ? {} : { senderCheck: verdict.senderCheck }),
  };
}

/**
 * Refuse a mailbox-changing call when the cached grant was taken with
 * `--readonly`. Graph would reject it anyway, but only after the request has
 * left the machine and with a generic 403; catching it here keeps the failure
 * local, immediate, and phrased in terms of the flag the user actually chose.
 *
 * Lives on guardMutation for the same reason --dry-run does: it is the one
 * path every remote mutation already goes through, so no command can forget it.
 */
async function assertWriteScope(ctx: CommandContext, spec: RequestSpec): Promise<void> {
  const status = await ctx.auth.status();
  // Not signed in at all is NOT this error's business — let the token fetch
  // raise NotAuthenticatedError with its own instructions.
  if (!status.signedIn || status.mode !== 'readonly') return;
  throw new InsufficientScopeError(`${spec.method} ${spec.url}`);
}

/**
 * The JSON twin of dryRun.ts's `renderRequest`. Both paths print a request the
 * user is about to send, so both must scrub it identically — this one used to
 * pass `body` through untouched, which meant `--dry-run --json` leaked what
 * `--dry-run` alone had already masked. Headers are dropped entirely rather
 * than redacted: the Authorization header is attached later, inside the send
 * path, so it is not in the spec at all and printing the rest adds nothing.
 */
function redactSpecForJson(spec: RequestSpec): RequestSpec {
  const { method, url, body } = spec;
  return {
    method,
    url,
    ...(body === undefined ? {} : { body: redactObject(body) }),
  };
}

/** Emit either JSON or a human rendering, chosen by --json. */
export function emit(
  ctx: CommandContext,
  data: unknown,
  human: () => string,
): void {
  ctx.out(ctx.json ? JSON.stringify(data, null, 2) : human());
}
