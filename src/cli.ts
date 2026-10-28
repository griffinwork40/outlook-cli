/**
 * `outlook` — the CLI entrypoint.
 *
 * Three jobs, and deliberately nothing else:
 *   1. Route argv[0] to a CommandSpec.
 *   2. Build the single CommandContext every command binds to. This is the
 *      ONLY place the real `fetch` and the real on-disk token cache are
 *      supplied — everything below takes them by injection, which is what
 *      lets the whole command layer be tested without a network.
 *   3. Own the one error -> exit-code boundary, so no command ever calls
 *      process.exit and every message is redacted exactly once on its way out.
 *
 * Exit codes: 0 ok | 1 unexpected | 2 usage | 3 auth/config (see errors.ts).
 */

import { createRequire } from 'node:module';

import { loadEnvFile } from './env.js';

// Load ~/.afk/config/afk.env before anything reads process.env — agent
// sessions may not have GRAPH_CLIENT_ID exported in the shell profile.
loadEnvFile();

import type { ParsedArgs } from './args.js';
import { parseArgs } from './args.js';
import { createAuthenticator } from './auth.js';
import { GraphClient } from './client.js';
import type { CommandContext, CommandSpec } from './command.js';
import { resolveConfig } from './config.js';
import { OutlookCliError, UsageError } from './errors.js';
import { assertKnownFlags } from './flagSpec.js';
import { redactText } from './redact.js';
import { createTokenStore } from './tokenStore.js';

import { attachmentsCommand } from './commands/attachments.js';
import { authCommand } from './commands/auth.js';
import { contactsCommand } from './commands/contacts.js';
import { draftCommand } from './commands/draft.js';
import {
  flagCommand,
  markReadCommand,
  markUnreadCommand,
  unflagCommand,
} from './commands/flags.js';
import { foldersCommand } from './commands/folders.js';
import { getCommand } from './commands/get.js';
import { listCommand } from './commands/list.js';
import { meCommand } from './commands/me.js';
import { deleteCommand, moveCommand, purgeCommand } from './commands/move.js';
import {
  forwardCommand,
  replyAllCommand,
  replyCommand,
} from './commands/reply.js';
import { rulesCommand } from './commands/rules.js';
import { searchCommand } from './commands/search.js';
import { sendCommand } from './commands/send.js';

const COMMANDS: readonly CommandSpec[] = [
  authCommand,
  meCommand,
  listCommand,
  getCommand,
  searchCommand,
  foldersCommand,
  attachmentsCommand,
  contactsCommand,
  sendCommand,
  replyCommand,
  replyAllCommand,
  forwardCommand,
  draftCommand,
  moveCommand,
  deleteCommand,
  purgeCommand,
  markReadCommand,
  markUnreadCommand,
  flagCommand,
  unflagCommand,
  rulesCommand,
];

const REGISTRY = new Map<string, CommandSpec>(
  COMMANDS.map((command) => [command.name, command]),
);

/** Help groupings. Names are resolved through REGISTRY, never duplicated. */
const GROUPS: readonly { title: string; names: readonly string[] }[] = [
  { title: 'Auth & identity', names: ['auth', 'me'] },
  {
    title: 'Read',
    names: ['list', 'get', 'search', 'folders', 'attachments'],
  },
  {
    title: 'People',
    names: ['contacts'],
  },
  {
    title: 'Send',
    names: ['send', 'reply', 'reply-all', 'forward', 'draft'],
  },
  {
    title: 'Organize',
    names: [
      'move',
      'delete',
      'purge',
      'mark-read',
      'mark-unread',
      'flag',
      'unflag',
      'rules',
    ],
  },
];

function packageVersion(): string {
  try {
    const pkg = createRequire(import.meta.url)('../package.json') as {
      version?: string;
    };
    return pkg.version ?? '0.0.0';
  } catch {
    return '0.0.0';
  }
}

function helpText(): string {
  const lines: string[] = [
    'outlook — Microsoft Graph Outlook Mail from the command line',
    '',
    'Usage: outlook <command> [options]',
    '',
  ];

  for (const group of GROUPS) {
    lines.push(`${group.title}:`);
    for (const name of group.names) {
      const command = REGISTRY.get(name);
      if (command === undefined) continue;
      const flag = command.mutating ? ' [--dry-run]' : '';
      lines.push(`  ${name.padEnd(13)} ${command.summary}${flag}`);
    }
    lines.push('');
  }

  lines.push(
    'Global options:',
    '  --json         Machine-readable output',
    '  --dry-run      On a mutating command: print the request, send nothing',
    '  -h, --help     This help, or `outlook <command> --help`',
    '  -V, --version  Print the version',
    '',
    'First run: export GRAPH_CLIENT_ID=<azure-app-id> && outlook auth login',
    'Setup walkthrough: SETUP.md',
    '',
  );

  return lines.join('\n');
}

/**
 * Wire the real dependencies. `getToken` is passed as a thunk so the client
 * asks the authenticator for a token per attempt — that is what makes a
 * mid-flight refresh transparent to every command.
 *
 * Note: config resolution requires GRAPH_CLIENT_ID for every command,
 * including `auth logout`. That is intentional — one code path, and the
 * MissingCredentialsError already names the variable and points at SETUP.md.
 */
function buildContext(args: ParsedArgs): CommandContext {
  const config = resolveConfig(process.env);
  const store = createTokenStore(config.tokenCachePath);
  const auth = createAuthenticator(config, store);
  const client = new GraphClient({
    getToken: () => auth.getAccessToken(),
  });

  return {
    args,
    json: args.bool('json'),
    dryRun: args.bool('dry-run'),
    client,
    auth,
    env: process.env,
    out: (line: string) => {
      process.stdout.write(`${line}\n`);
    },
    err: (line: string) => {
      process.stderr.write(`${line}\n`);
    },
  };
}

async function run(argv: string[]): Promise<number> {
  const first = argv[0];

  if (first === undefined || first === 'help' || first === '--help' || first === '-h') {
    process.stdout.write(helpText());
    return 0;
  }
  if (first === 'version' || first === '--version' || first === '-V') {
    process.stdout.write(`${packageVersion()}\n`);
    return 0;
  }

  const command = REGISTRY.get(first);
  if (command === undefined) {
    throw new UsageError(
      `Unknown command: ${first}\n\nRun 'outlook help' for the command list.`,
    );
  }

  const rest = argv.slice(1);
  const args = parseArgs(rest);

  if (args.bool('help') || args.bool('h')) {
    process.stdout.write(`${command.summary}\n\n${command.usage}\n`);
    return 0;
  }

  // Fail closed on a misspelled flag BEFORE any context is built or any
  // handler runs. A typo'd --dry-run must never reach a mutating command.
  assertKnownFlags(command.name, args.names());

  const ctx = buildContext(args);

  // A no-op --dry-run should say so rather than imply something was withheld.
  if (ctx.dryRun && !command.mutating) {
    ctx.err(
      `note: '${command.name}' does not change anything, so --dry-run has no effect.`,
    );
  }

  await command.handler(ctx);
  return 0;
}

function reportFailure(error: unknown): number {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`${redactText(message)}\n`);
  return error instanceof OutlookCliError ? error.exitCode : 1;
}

// Set process.exitCode rather than calling process.exit(), so buffered
// stdout is never truncated mid-write on a piped or redirected run.
try {
  process.exitCode = await run(process.argv.slice(2));
} catch (error) {
  process.exitCode = reportFailure(error);
}
