/**
 * Tests for `outlook rules list` and helpers. See rules-create.test.ts for
 * the `create` subcommand tests.
 */

import { describe, expect, it, vi } from 'vitest';
import { UsageError } from '../errors.js';
import type { GraphCollection } from '../types.js';
import {
  rulesCommand,
  summariseConditions,
  summariseActions,
  type MessageRule,
} from './rules.js';
import { makeCtx, text } from './testUtils.js';

const SAMPLE_RULES: GraphCollection<MessageRule> = {
  value: [
    {
      id: 'rule-1',
      displayName: 'Move newsletters',
      sequence: 1,
      isEnabled: true,
      isReadOnly: false,
      conditions: {
        fromAddresses: [
          { emailAddress: { address: 'news@example.com' } },
        ],
      },
      actions: { moveToFolder: 'AAMkNewsletter' },
    },
    {
      id: 'rule-2',
      displayName: 'Flag from boss',
      sequence: 2,
      isEnabled: false,
      isReadOnly: true,
      conditions: {
        fromAddresses: [
          { emailAddress: { address: 'boss@contoso.com', name: 'Boss' } },
        ],
      },
      actions: { markImportance: 'high', stopProcessingRules: true },
    },
  ],
};

describe('rules list', () => {
  it('GETs /me/mailFolders/inbox/messageRules and renders human output', async () => {
    const { ctx, client, out } = makeCtx(['list']);
    const get = vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    expect(get).toHaveBeenCalledTimes(1);
    expect(get.mock.calls[0]?.[0]).toBe('/me/mailFolders/inbox/messageRules');
    expect(text(out)).toContain('Move newsletters');
    expect(text(out)).toContain('Flag from boss');
  });

  it('defaults to "list" when no subcommand is given', async () => {
    const { ctx, client } = makeCtx([]);
    const get = vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    expect(get).toHaveBeenCalledTimes(1);
  });

  it('--json emits the raw rules array', async () => {
    const { ctx, client, out } = makeCtx(['list', '--json']);
    vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(SAMPLE_RULES.value);
  });

  it('without --json, output is human text, not a JSON blob', async () => {
    const { ctx, client, out } = makeCtx(['list']);
    vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('renders "(no rules)" when the list is empty', async () => {
    const { ctx, client, out } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue({ value: [] });

    await rulesCommand.handler(ctx);

    expect(text(out)).toContain('(no rules)');
  });

  it('shows enabled status for each rule', async () => {
    const { ctx, client, out } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    const output = text(out);
    expect(output).toContain('yes');
    expect(output).toContain('no');
  });

  it('throws UsageError for an unknown subcommand with the expected message', async () => {
    const { ctx } = makeCtx(['bogus']);

    await expect(rulesCommand.handler(ctx)).rejects.toThrow(
      /Expected "list" or "create"/,
    );
  });

  it('never touches the network directly', async () => {
    const { ctx, client, fetchImpl } = makeCtx([]);
    vi.spyOn(client, 'get').mockResolvedValue(SAMPLE_RULES);

    await rulesCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});

describe('summariseConditions', () => {
  it('returns "(any)" for undefined or empty conditions', () => {
    expect(summariseConditions(undefined)).toBe('(any)');
    expect(summariseConditions({})).toBe('(any)');
  });

  it('summarises fromAddresses', () => {
    const result = summariseConditions({
      fromAddresses: [{ emailAddress: { address: 'a@b.com' } }],
    });
    expect(result).toBe('from: a@b.com');
  });

  it('summarises subjectContains', () => {
    const result = summariseConditions({ subjectContains: ['urgent', 'ASAP'] });
    expect(result).toBe('subject contains: urgent, ASAP');
  });

  it('summarises hasAttachments', () => {
    expect(summariseConditions({ hasAttachments: true })).toBe('has attachments');
  });

  it('summarises senderContains', () => {
    const result = summariseConditions({ senderContains: ['@example.com'] });
    expect(result).toBe('sender contains: @example.com');
  });
});

describe('summariseActions', () => {
  it('returns "(none)" for undefined or empty actions', () => {
    expect(summariseActions(undefined)).toBe('(none)');
    expect(summariseActions({})).toBe('(none)');
  });

  it('summarises moveToFolder', () => {
    expect(summariseActions({ moveToFolder: 'AAMk...' })).toBe('move to [AAMk...]');
  });

  it('summarises delete', () => {
    expect(summariseActions({ delete: true })).toBe('delete');
  });

  it('summarises multiple actions', () => {
    const result = summariseActions({
      markAsRead: true,
      stopProcessingRules: true,
    });
    expect(result).toBe('mark read; stop processing');
  });
});
