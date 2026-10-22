/**
 * The set of flags each command accepts, and the rejection of everything else.
 *
 * Why this exists: the argv parser is deliberately schema-less, so it records
 * whatever the user typed under whatever name they typed. Nothing downstream
 * ever asked "is that a real flag?", which meant a typo was indistinguishable
 * from an omission. The consequence was not cosmetic:
 *
 *     outlook send --to a@b.com --subject S --body B --dryrun
 *
 * parses `dryrun`, nobody reads `dryrun`, `ctx.dryRun` is therefore false, the
 * guardMutation chokepoint is skipped, and the mail is really sent. One missing
 * hyphen turned "show me what would go out" into "send it now." Fail closed.
 *
 * The map lives in ONE file rather than being colocated on each CommandSpec so
 * the whole authorized surface is reviewable at a glance, and so a single test
 * can assert that every registered command has an entry (flagSpec.test.ts).
 * A command with no flags of its own must still be listed, with an empty array
 * — absence is a bug, not "allow everything".
 */

import { UsageError } from './errors.js';

/** Accepted on every command, so each entry below lists only its own flags. */
export const GLOBAL_FLAGS: readonly string[] = ['json', 'dry-run', 'help', 'h'];

/** Shared by the four compose surfaces (send/reply-family/draft) via compose.ts. */
const COMPOSE_FLAGS: readonly string[] = [
  'to',
  'cc',
  'bcc',
  'subject',
  'body',
  'body-file',
  'html',
  'attach',
  'signature',
];

/**
 * command name -> flags it accepts, beyond GLOBAL_FLAGS.
 *
 * Note on `--no-x`: the parser strips the `no-` prefix and records the bare
 * name, so `--no-save-sent` validates against `save-sent`. Only the bare name
 * belongs here.
 */
export const COMMAND_FLAGS: Readonly<Record<string, readonly string[]>> = {
  auth: ['readonly'],
  me: [],
  list: ['folder', 'limit', 'all', 'unread'],
  get: ['html'],
  search: ['limit', 'all', 'folder'],
  folders: ['parent', 'all', 'name'],
  attachments: ['all', 'out', 'force'],
  send: [...COMPOSE_FLAGS, 'save-sent'],
  reply: ['body', 'signature'],
  'reply-all': ['body', 'signature'],
  forward: ['body', 'to', 'signature'],
  draft: [...COMPOSE_FLAGS, 'limit', 'all'],
  move: ['folder'],
  delete: [],
  purge: ['yes'],
  'mark-read': [],
  'mark-unread': [],
  flag: [],
  unflag: [],
  rules: ['name', 'folder', 'from', 'sender-contains', 'mark-read', 'stop-processing', 'enabled', 'sequence'],
};

/** Levenshtein, capped — only used to build a "did you mean" hint. */
function editDistance(a: string, b: string): number {
  const rows = a.length + 1;
  const cols = b.length + 1;
  let prev: number[] = Array.from({ length: cols }, (_, j) => j);

  for (let i = 1; i < rows; i++) {
    const curr: number[] = [i];
    for (let j = 1; j < cols; j++) {
      const substitution = (prev[j - 1] ?? 0) + (a[i - 1] === b[j - 1] ? 0 : 1);
      const insertion = (curr[j - 1] ?? 0) + 1;
      const deletion = (prev[j] ?? 0) + 1;
      curr[j] = Math.min(substitution, insertion, deletion);
    }
    prev = curr;
  }
  return prev[cols - 1] ?? Math.max(a.length, b.length);
}

/** Closest allowed flag within a small edit distance, for the error hint. */
function suggest(unknown: string, allowed: readonly string[]): string | undefined {
  let best: string | undefined;
  let bestScore = Number.POSITIVE_INFINITY;

  for (const candidate of allowed) {
    const score = editDistance(unknown.toLowerCase(), candidate.toLowerCase());
    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  // Distance 3 catches `dryrun` -> `dry-run` (1) and `subjekt` -> `subject` (1)
  // without proposing nonsense for a genuinely unrelated token.
  const ceiling = Math.min(3, Math.max(1, Math.floor(unknown.length / 2)));
  return best !== undefined && bestScore <= ceiling ? best : undefined;
}

/**
 * Throw if the invocation carries a flag the command does not define.
 *
 * @param commandName registered command name (used to look up the allowlist)
 * @param used        every flag name the parser actually saw
 */
export function assertKnownFlags(commandName: string, used: readonly string[]): void {
  const own = COMMAND_FLAGS[commandName];
  if (own === undefined) {
    // An unregistered command reaching here means the registry and this map
    // drifted apart. Fail closed rather than silently allowing everything.
    throw new UsageError(
      `Internal: no flag specification registered for command '${commandName}'.`,
    );
  }

  const allowed = [...GLOBAL_FLAGS, ...own];
  const allowedSet = new Set(allowed);
  const unknown = used.filter((name) => !allowedSet.has(name));
  if (unknown.length === 0) return;

  const first = unknown[0] ?? '';
  const hint = suggest(first, allowed);
  const lines = [
    `Unknown flag for '${commandName}': --${unknown.map((u) => u).join(', --')}`,
  ];
  if (hint !== undefined) lines.push(`Did you mean --${hint}?`);
  lines.push(
    '',
    `Accepted: ${allowed.map((f) => `--${f}`).join(' ')}`,
    `Run 'outlook ${commandName} --help' for usage.`,
  );

  throw new UsageError(lines.join('\n'));
}
