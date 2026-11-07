/**
 * Local token cache: a single JSON file, chmod 0600. A corrupt or unreadable
 * cache must never crash the CLI — it just means "sign in again," so `load`
 * swallows every failure mode (missing file, bad JSON, wrong shape) and
 * returns `null` rather than throwing.
 */

import { randomBytes } from 'node:crypto';
import { chmod, mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { isScopeArray } from './scopes.js';
import type { TokenSet } from './types.js';

/** Owner read/write only — the file holds an access + refresh token. */
const FILE_MODE = 0o600;

/**
 * Owner-only on the containing directory too. The file itself is already 0600,
 * so this is not what protects the secret — it stops a world-readable directory
 * from advertising that a token cache exists and when it was last written.
 */
const DIR_MODE = 0o700;

/**
 * chmod cannot express "owner only" through the Win32 ACL model, so on Windows
 * the 0600 above is decorative and the token file inherits directory
 * permissions. Callers are told once, loudly, rather than being left with a
 * false sense of protection. See docs/security-review-2026-08-07.md (#4).
 */
export const FILE_MODE_IS_ENFORCED = process.platform !== 'win32';

export interface TokenStore {
  readonly path: string;
  load(): Promise<TokenSet | null>;
  save(tokens: TokenSet): Promise<void>;
  clear(): Promise<void>;
}

export function createTokenStore(path: string): TokenStore {
  return {
    path,

    async load(): Promise<TokenSet | null> {
      try {
        const raw = await readFile(path, 'utf8');
        const parsed: unknown = JSON.parse(raw);
        return isTokenSet(parsed) ? parsed : null;
      } catch {
        // ENOENT, EACCES, malformed JSON, a directory where a file should
        // be — all of it collapses to "no usable cache." Never throw.
        return null;
      }
    },

    /**
     * Atomic: write a fresh temp file, lock it down, then rename over the
     * target. rename(2) is atomic within a directory, so a reader either sees
     * the whole old token set or the whole new one — never a half-written
     * file. The previous in-place writeFile could leave a truncated cache
     * behind if the process died mid-write, or interleave two concurrent
     * refreshes into corrupt JSON, either of which forces a re-login.
     *
     * The temp file carries the restrictive mode from creation, so the secret
     * is never briefly world-readable in the window before chmod lands.
     *
     * The temp NAME is random and the open is exclusive. The previous
     * `.${pid}-${Date.now()}.token.tmp` was guessable, and writeFile's default
     * O_CREAT|O_TRUNC follows a symlink — so a link pre-planted at that path
     * would have received the token at whatever destination it pointed to.
     * 'wx' (O_CREAT|O_EXCL) refuses to open anything that already exists.
     */
    async save(tokens: TokenSet): Promise<void> {
      const dir = dirname(path);
      await mkdir(dir, { recursive: true, mode: DIR_MODE });

      const tmp = join(dir, `.${randomBytes(8).toString('hex')}.token.tmp`);
      try {
        await writeFile(tmp, JSON.stringify(tokens, null, 2), {
          mode: FILE_MODE,
          flag: 'wx',
        });
        // `mode` on writeFile only applies at creation; chmod makes it
        // explicit. On Windows both are advisory — see FILE_MODE_IS_ENFORCED.
        await chmod(tmp, FILE_MODE);
        await rename(tmp, path);
      } catch (error) {
        // EEXIST means the path was already occupied by something this
        // function did not create. Removing it would be deleting a stranger's
        // file on the way out of a failure — leave it alone.
        if ((error as NodeJS.ErrnoException).code !== 'EEXIST') {
          await rm(tmp, { force: true }).catch(() => undefined);
        }
        throw error;
      }
    },

    async clear(): Promise<void> {
      // force: true swallows ENOENT — clearing an already-absent cache is a
      // no-op, not an error.
      await rm(path, { force: true });
    },
  };
}

/** Runtime shape guard — a corrupt cache must degrade to `null`, not a crash. */
function isTokenSet(value: unknown): value is TokenSet {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;

  if (typeof v['accessToken'] !== 'string' || v['accessToken'].length === 0) return false;
  if (typeof v['expiresAt'] !== 'number' || !Number.isFinite(v['expiresAt'])) return false;

  if (v['refreshToken'] !== undefined && typeof v['refreshToken'] !== 'string') return false;
  if (v['scope'] !== undefined && typeof v['scope'] !== 'string') return false;
  if (v['tokenType'] !== undefined && typeof v['tokenType'] !== 'string') return false;
  if (v['account'] !== undefined && typeof v['account'] !== 'string') return false;
  // An unrecognized mode must not silently read as 'full' (that would be a
  // privilege upgrade on a corrupt cache) — reject the whole cache instead.
  if (v['mode'] !== undefined && v['mode'] !== 'readonly' && v['mode'] !== 'full') {
    return false;
  }
  // Same reasoning: a malformed custom scope set must not fall back to the
  // default (wider) scopes on the next refresh.
  if (v['grantScopes'] !== undefined && !isScopeArray(v['grantScopes'])) return false;

  return true;
}
