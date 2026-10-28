/**
 * Tests for `outlook contacts`. Uses the same idiom as list.test.ts:
 * real GraphClient from makeCtx, network boundary stubbed via
 * vi.spyOn(client, 'request') for collectItems paths and
 * vi.spyOn(client, 'get') for the single-resource `get` subcommand.
 */

import { describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../client.js';
import { UsageError } from '../errors.js';
import { GRAPH_BASE, type Contact, type GraphCollection, type RequestSpec } from '../types.js';
import { contactsCommand } from './contacts.js';
import { makeCtx, text } from './testUtils.js';

function stubPages(client: GraphClient, ...pages: GraphCollection<Contact>[]) {
  const spy = vi.spyOn(client, 'request');
  for (const page of pages) spy.mockResolvedValueOnce(page as never);
  return spy;
}

/** Relative Graph path of a captured RequestSpec (spec.url is absolute). */
function pathOf(spy: { mock: { calls: unknown[][] } }, call = 0): string {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  if (spec === undefined) return '';
  const url = spec.url.startsWith(GRAPH_BASE) ? spec.url.slice(GRAPH_BASE.length) : spec.url;
  return url.split('?')[0] ?? '';
}

function queryOf(spy: { mock: { calls: unknown[][] } }, call = 0): URLSearchParams {
  const spec = spy.mock.calls[call]?.[0] as RequestSpec | undefined;
  return new URL(spec?.url ?? GRAPH_BASE).searchParams;
}

function sampleContact(overrides: Partial<Contact> = {}): Contact {
  return {
    id: 'AAMk-contact-0001',
    displayName: 'Alice Smith',
    emailAddresses: [{ address: 'alice@example.com', name: 'Alice Smith' }],
    mobilePhone: '+1-555-0100',
    businessPhones: ['+1-555-0101'],
    homePhones: [],
    companyName: 'Contoso',
    jobTitle: 'Engineer',
    ...overrides,
  };
}

describe('contacts', () => {
  it('contacts list returns formatted list', async () => {
    const { ctx, client, out } = makeCtx(['list']);
    stubPages(client, { value: [sampleContact()] });

    await contactsCommand.handler(ctx);

    expect(text(out)).toContain('Alice Smith');
    expect(text(out)).toContain('Contoso');
    expect(() => JSON.parse(text(out))).toThrow();
  });

  it('contacts list --json returns JSON array', async () => {
    const { ctx, client, out } = makeCtx(['list', '--json']);
    const contacts = [sampleContact()];
    stubPages(client, { value: contacts });

    await contactsCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(contacts);
  });

  it('contacts list --top 5 passes $top=5', async () => {
    const { ctx, client } = makeCtx(['list', '--top', '5']);
    const req = stubPages(client, { value: [sampleContact()] });

    await contactsCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(queryOf(req).get('$top')).toBe('5');
  });

  it('contacts list hits /me/contacts with correct defaults', async () => {
    const { ctx, client } = makeCtx(['list']);
    const req = stubPages(client, { value: [] });

    await contactsCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/contacts');
    expect(queryOf(req).get('$top')).toBe('25');
    expect(queryOf(req).get('$orderby')).toBe('displayName asc');
    expect(queryOf(req).get('$select')).toBe(
      'id,displayName,emailAddresses,mobilePhone,businessPhones,homePhones,companyName,jobTitle',
    );
  });

  it('contacts get <id> fetches and renders single contact', async () => {
    const { ctx, client, out } = makeCtx(['get', 'AAMk-contact-0001']);
    const contact = sampleContact();
    const get = vi.spyOn(client, 'get').mockResolvedValue(contact);

    await contactsCommand.handler(ctx);

    expect(get).toHaveBeenCalledTimes(1);
    expect(get.mock.calls[0]?.[0]).toBe('/me/contacts/AAMk-contact-0001');
    expect(text(out)).toContain('Name: Alice Smith');
    expect(text(out)).toContain('alice@example.com');
    expect(text(out)).toContain('+1-555-0100');
  });

  it('contacts get <id> --json returns JSON', async () => {
    const { ctx, client, out } = makeCtx(['get', 'AAMk-contact-0001', '--json']);
    const contact = sampleContact();
    vi.spyOn(client, 'get').mockResolvedValue(contact);

    await contactsCommand.handler(ctx);

    expect(JSON.parse(text(out))).toEqual(contact);
  });

  it('contacts search <query> uses $filter with startswith', async () => {
    const { ctx, client } = makeCtx(['search', 'Ali']);
    const req = stubPages(client, { value: [sampleContact()] });

    await contactsCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/contacts');
    expect(queryOf(req).get('$filter')).toBe("startswith(displayName,'Ali')");
  });

  it('contacts search escapes single quotes in query', async () => {
    const { ctx, client } = makeCtx(['search', "O'Brien"]);
    const req = stubPages(client, { value: [] });

    await contactsCommand.handler(ctx);

    expect(queryOf(req).get('$filter')).toBe("startswith(displayName,'O''Brien')");
  });

  it('contacts with no subcommand defaults to list', async () => {
    const { ctx, client } = makeCtx([]);
    const req = stubPages(client, { value: [sampleContact()] });

    await contactsCommand.handler(ctx);

    expect(req).toHaveBeenCalledTimes(1);
    expect(pathOf(req)).toBe('/me/contacts');
  });

  it('unknown subcommand throws UsageError', async () => {
    const { ctx, client } = makeCtx(['bogus']);
    const req = vi.spyOn(client, 'request');

    await expect(contactsCommand.handler(ctx)).rejects.toBeInstanceOf(UsageError);
    expect(req).toHaveBeenCalledTimes(0);
  });

  it('never touches the network directly — only through stubbed client methods', async () => {
    const { ctx, client, fetchImpl } = makeCtx(['list']);
    stubPages(client, { value: [sampleContact()] });

    await contactsCommand.handler(ctx);

    expect(fetchImpl).toHaveBeenCalledTimes(0);
  });
});
