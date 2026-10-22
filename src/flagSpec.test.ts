/**
 * Tests for flagSpec.ts — the per-command flag allowlist that fail-closes on
 * typos (see the file header there for the `--dryrun` incident this exists
 * to prevent).
 *
 * The command modules are imported directly, NOT via cli.ts: cli.ts runs
 * top-level routing/registry-building code on import and does not export its
 * COMMANDS array, so importing it here would be both unnecessary and risky.
 * Each command module already exports its CommandSpec(s) individually (as
 * every sibling command test file relies on), so the registry is rebuilt
 * locally from those exports for the completeness assertions below.
 */

import { describe, expect, it } from 'vitest';
import { assertKnownFlags, COMMAND_FLAGS, GLOBAL_FLAGS } from './flagSpec.js';
import { UsageError } from './errors.js';

import { attachmentsCommand } from './commands/attachments.js';
import { authCommand } from './commands/auth.js';
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
import { searchCommand } from './commands/search.js';
import { sendCommand } from './commands/send.js';
import { rulesCommand } from './commands/rules.js';

/**
 * Rebuilt independently of cli.ts's (unexported) COMMANDS array, from each
 * command module's own exports — the same exports cli.ts itself imports.
 */
const REGISTERED_COMMANDS = [
  authCommand,
  meCommand,
  listCommand,
  getCommand,
  searchCommand,
  foldersCommand,
  attachmentsCommand,
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
] as const;

describe('assertKnownFlags — known flags pass', () => {
  it("accepts flags that are in the command's own list (list: folder, limit)", () => {
    expect(() => assertKnownFlags('list', ['folder', 'limit'])).not.toThrow();
  });

  it('accepts json, dry-run, help, and h individually on a representative command', () => {
    expect(() => assertKnownFlags('list', ['json'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['dry-run'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['help'])).not.toThrow();
    expect(() => assertKnownFlags('list', ['h'])).not.toThrow();
  });

  it('accepts GLOBAL_FLAGS on every registered command', () => {
    for (const command of REGISTERED_COMMANDS) {
      expect(() => assertKnownFlags(command.name, [...GLOBAL_FLAGS])).not.toThrow();
    }
  });
});

describe('assertKnownFlags — unknown flags throw', () => {
  it('throws UsageError for a flag not in the allowlist', () => {
    expect(() => assertKnownFlags('list', ['bogus'])).toThrow(UsageError);
  });

  it('THE CRITICAL ONE: rejects the --dryrun typo and suggests --dry-run — this is the typo that used to send real mail', () => {
    expect(() => assertKnownFlags('send', ['dryrun'])).toThrow(UsageError);
    expect(() => assertKnownFlags('send', ['dryrun'])).toThrow(/Did you mean --dry-run\?/);
  });

  it("throws for an unknown flag on a command with an empty own-flag list ('me')", () => {
    expect(() => assertKnownFlags('me', ['bogus'])).toThrow(UsageError);
  });
});

describe('assertKnownFlags — unregistered command name (fail closed)', () => {
  it('throws UsageError rather than silently allowing everything', () => {
    expect(() => assertKnownFlags('not-a-real-command', [])).toThrow(UsageError);
    expect(() => assertKnownFlags('not-a-real-command', [])).toThrow(
      /no flag specification registered/,
    );
  });
});

describe('COMMAND_FLAGS — completeness against the real command registry', () => {
  it('has an entry for every registered command', () => {
    for (const command of REGISTERED_COMMANDS) {
      expect(COMMAND_FLAGS[command.name]).toBeDefined();
    }
  });

  it('has no stale entries that do not correspond to a registered command', () => {
    const registeredNames = new Set(REGISTERED_COMMANDS.map((command) => command.name));
    for (const name of Object.keys(COMMAND_FLAGS)) {
      expect(registeredNames.has(name)).toBe(true);
    }
  });
});

describe('assertKnownFlags — rules command (end-to-end flag validation)', () => {
  it('accepts all flags used by rules list (json only)', () => {
    expect(() => assertKnownFlags('rules', ['json'])).not.toThrow();
    expect(() => assertKnownFlags('rules', ['dry-run'])).not.toThrow();
  });

  it('accepts all flags used by rules create', () => {
    const createFlags = [
      'name', 'folder', 'from', 'sender-contains',
      'mark-read', 'stop-processing', 'enabled', 'sequence', 'dry-run',
    ];
    expect(() => assertKnownFlags('rules', createFlags)).not.toThrow();
  });

  it('rejects a typo like --dispaly-name and suggests --name', () => {
    expect(() => assertKnownFlags('rules', ['dispaly-name'])).toThrow(UsageError);
  });

  it('rejects a completely unknown flag', () => {
    expect(() => assertKnownFlags('rules', ['bogus'])).toThrow(UsageError);
  });
});
