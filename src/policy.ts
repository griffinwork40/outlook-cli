/**
 * The OUTLOOK_POLICY file: loading and validation.
 *
 * A policy confines what an unattended caller may change in the mailbox. It is
 * read fresh on every guarded mutation and NEVER consulted for reads. Every
 * failure here is reported as `{ ok: false, reason }` rather than thrown,
 * because the caller's answer to any of them is the same: refuse (fail closed).
 * A policy that cannot be read is not "no policy".
 *
 * Example:
 *   {
 *     "allowCommands": ["move", "mark-read"],
 *     "move": {
 *       "allowedDestinations": ["<review-folder-id>", "archive"],
 *       "denySenders": {
 *         "file": "protected_senders.json",
 *         "maxAgeHours": 48,
 *         "destinations": ["<review-folder-id>"]
 *       }
 *     }
 *   }
 */

import { readFile, stat } from 'node:fs/promises';
import { isWellKnownFolder } from './types.js';

export const POLICY_ENV = 'OUTLOOK_POLICY';

export interface DenySendersRule {
  /** JSON file: an array of addresses, or `{ addresses: [...], built_at: ISO }`. Relative to the policy file. */
  file: string;
  /** Older than this (by built_at, else file mtime) and the list is not trusted. */
  maxAgeHours: number;
  /** Moves INTO these folders are checked against the list. */
  destinations: string[];
}

export interface MutationPolicy {
  /** Command names allowed to mutate, e.g. "move", "draft create". A bare "draft" allows every draft subcommand. */
  allowCommands: string[];
  move?: {
    allowedDestinations?: string[];
    denySenders?: DenySendersRule;
  };
}

export type Loaded<T> = { ok: true; value: T } | { ok: false; reason: string };

const TOP_KEYS = new Set(['allowCommands', 'move']);
const MOVE_KEYS = new Set(['allowedDestinations', 'denySenders']);
const DENY_KEYS = new Set(['file', 'maxAgeHours', 'destinations']);

/** A UTF-8 BOM (common in files written by Windows tools) is not JSON. */
function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

function isStringList(v: unknown, allowEmpty: boolean): v is string[] {
  return (
    Array.isArray(v) &&
    (allowEmpty || v.length > 0) &&
    v.every((s) => typeof s === 'string' && s.trim().length > 0)
  );
}

/**
 * Unknown keys are rejected (keys starting with `$`, e.g. `$comment`, are
 * allowed). A typo such as `allowedDestination` must not silently mean
 * "no destination restriction".
 */
function unknownKey(obj: Record<string, unknown>, known: Set<string>): string | undefined {
  return Object.keys(obj).find((k) => !known.has(k) && !k.startsWith('$'));
}

/** Returns a description of the first problem, or undefined when valid. */
export function validatePolicy(v: unknown): string | undefined {
  if (!isRecord(v)) return 'top level must be a JSON object';
  const extra = unknownKey(v, TOP_KEYS);
  if (extra !== undefined) return `unknown key "${extra}"`;
  if (!isStringList(v['allowCommands'], true)) {
    return '"allowCommands" must be an array of command names (it may be empty)';
  }

  const move = v['move'];
  if (move === undefined) return undefined;
  if (!isRecord(move)) return '"move" must be an object';
  const extraMove = unknownKey(move, MOVE_KEYS);
  if (extraMove !== undefined) return `unknown key "move.${extraMove}"`;

  const allowed = move['allowedDestinations'];
  if (allowed !== undefined && !isStringList(allowed, true)) {
    return '"move.allowedDestinations" must be an array of folder ids or well-known names';
  }

  const deny = move['denySenders'];
  if (deny === undefined) return undefined;
  if (!isRecord(deny)) return '"move.denySenders" must be an object';
  const extraDeny = unknownKey(deny, DENY_KEYS);
  if (extraDeny !== undefined) return `unknown key "move.denySenders.${extraDeny}"`;
  if (typeof deny['file'] !== 'string' || deny['file'].trim() === '') {
    return '"move.denySenders.file" must be a path';
  }
  const denyFile = deny['file'] as string;
  if (
    denyFile.startsWith('/') ||
    denyFile.includes('\\') ||
    denyFile.split('/').some((seg) => seg === '..')
  ) {
    return '"move.denySenders.file" must be a relative path with no \'..\' segments or backslashes';
  }
  const maxAge = deny['maxAgeHours'];
  if (typeof maxAge !== 'number' || !Number.isFinite(maxAge) || maxAge <= 0) {
    return '"move.denySenders.maxAgeHours" must be a positive number';
  }
  if (!isStringList(deny['destinations'], false)) {
    return '"move.denySenders.destinations" must be a non-empty array of folder ids';
  }
  return undefined;
}

export async function loadPolicy(path: string): Promise<Loaded<MutationPolicy>> {
  let raw: string;
  try {
    raw = await readFile(path, 'utf8');
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code ?? 'read error';
    return { ok: false, reason: `cannot read policy file ${path} (${code})` };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(stripBom(raw));
  } catch {
    return { ok: false, reason: `policy file ${path} is not valid JSON` };
  }

  const problem = validatePolicy(parsed);
  if (problem !== undefined) {
    return { ok: false, reason: `policy file ${path} is invalid: ${problem}` };
  }
  return { ok: true, value: parsed as MutationPolicy };
}

/**
 * Load the sender deny list. Missing, unparseable, wrongly shaped, or stale
 * all fail: a list that cannot be trusted to be current cannot vouch that a
 * sender is NOT on it.
 */
export async function loadDenyList(
  file: string,
  maxAgeHours: number,
  nowMs: number,
): Promise<Loaded<Set<string>>> {
  let raw: string;
  try {
    raw = await readFile(file, 'utf8');
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code ?? 'read error';
    return { ok: false, reason: `cannot read deny list ${file} (${code})` };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(stripBom(raw));
  } catch {
    return { ok: false, reason: `deny list ${file} is not valid JSON` };
  }

  let addresses: unknown;
  let builtAt: unknown;
  if (Array.isArray(parsed)) {
    addresses = parsed;
  } else if (isRecord(parsed)) {
    addresses = parsed['addresses'];
    builtAt = parsed['built_at'];
  }
  // An empty addresses array is rejected (fail-closed): an operator who writes
  // `{"addresses":[]}` has not configured a list that can vouch any sender is
  // absent; treat it the same as a missing or corrupt file. If the intent is
  // "allow all moves to this destination", remove the denySenders rule entirely.
  if (!isStringList(addresses, false)) {
    return {
      ok: false,
      reason: `deny list ${file} must be a non-empty array of addresses or an object with a non-empty "addresses" array (an empty list is rejected — omit denySenders to allow all senders)`,
    };
  }

  let stampMs: number;
  if (builtAt !== undefined) {
    stampMs = typeof builtAt === 'string' ? Date.parse(builtAt) : Number.NaN;
    if (Number.isNaN(stampMs)) {
      return { ok: false, reason: `deny list ${file} has an unparseable built_at` };
    }
  } else {
    try {
      stampMs = (await stat(file)).mtimeMs;
    } catch {
      return { ok: false, reason: `cannot stat deny list ${file}` };
    }
  }

  const ageHours = (nowMs - stampMs) / 3_600_000;
  if (ageHours > maxAgeHours) {
    return {
      ok: false,
      reason: `deny list ${file} is ${ageHours.toFixed(1)}h old (maxAgeHours ${maxAgeHours}); rebuild it`,
    };
  }

  return { ok: true, value: new Set(addresses.map((a) => a.trim().toLowerCase())) };
}

/**
 * Folder identity for policy matching: well-known names (`inbox`, `archive`)
 * compare case-insensitively, raw folder ids are opaque and compare exactly.
 * A well-known name and the id of that same folder do NOT match — list the
 * form the caller will actually pass.
 */
export function sameFolder(a: string, b: string): boolean {
  if (a === b) return true;
  return isWellKnownFolder(a) && isWellKnownFolder(b) && a.toLowerCase() === b.toLowerCase();
}
