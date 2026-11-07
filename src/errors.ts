/**
 * Error taxonomy. Every error carries an exit code so cli.ts can map
 * failures to shell semantics without re-inspecting types.
 *
 * Exit codes:
 *   1 — runtime/API failure
 *   2 — usage error (bad flags, missing args)
 *   3 — missing credentials / not authenticated
 *   4 — refused by the OUTLOOK_POLICY mutation policy
 */

export class OutlookCliError extends Error {
  readonly exitCode: number;

  constructor(message: string, exitCode = 1) {
    super(message);
    this.name = 'OutlookCliError';
    this.exitCode = exitCode;
  }
}

/** A required env var is absent. Names the exact var and points at SETUP.md. */
export class MissingCredentialsError extends OutlookCliError {
  readonly missing: readonly string[];

  constructor(missing: readonly string[]) {
    const list = missing.join(', ');
    super(
      `Missing required environment variable(s): ${list}\n` +
        `\nGRAPH_CLIENT_ID comes from a one-time Azure app registration.\n` +
        `See SETUP.md for the 5-step click-path, then export it:\n` +
        `  export GRAPH_CLIENT_ID=<application-client-id>`,
      3,
    );
    this.name = 'MissingCredentialsError';
    this.missing = missing;
  }
}

/** No cached token, or the cached token is unusable and cannot be refreshed. */
export class NotAuthenticatedError extends OutlookCliError {
  constructor(detail?: string) {
    super(
      `Not signed in${detail ? `: ${detail}` : '.'}\n` +
        `\nRun:  outlook auth login`,
      3,
    );
    this.name = 'NotAuthenticatedError';
  }
}

/**
 * The cached grant is read-only and the command needs to change the mailbox.
 * Caught locally, before the request goes out, so a least-privilege session
 * fails with an actionable message instead of an opaque Graph 403.
 */
export class InsufficientScopeError extends OutlookCliError {
  constructor(action: string) {
    super(
      `This session was signed in with --readonly, so it cannot ${action}.\n` +
        `\nRe-authorize with full access:  outlook auth login`,
      3,
    );
    this.name = 'InsufficientScopeError';
  }
}

/**
 * A mutation refused by the OUTLOOK_POLICY file (policy.ts). Its own exit code
 * so an unattended caller can tell "the guard said no" from a Graph failure
 * and stop instead of retrying. `rule` names the policy clause that fired.
 */
export class PolicyRefusedError extends OutlookCliError {
  readonly rule: string;

  constructor(rule: string, reason: string) {
    super(`Refused by mutation policy (${rule}): ${reason}`, 4);
    this.name = 'PolicyRefusedError';
    this.rule = rule;
  }
}

/** Bad flags or arguments. */
export class UsageError extends OutlookCliError {
  constructor(message: string) {
    super(message, 2);
    this.name = 'UsageError';
  }
}

/** A non-2xx response from Graph or the identity platform. */
export class GraphError extends OutlookCliError {
  readonly status: number;
  readonly code: string;
  readonly graphMessage: string;
  readonly requestId?: string;
  readonly retryAfter?: number;

  constructor(params: {
    status: number;
    code: string;
    message: string;
    requestId?: string;
    retryAfter?: number;
  }) {
    const hint = GRAPH_ERROR_HINTS[params.code];
    super(
      `Graph API error ${params.status} (${params.code}): ${params.message}` +
        (hint ? `\n  → ${hint}` : '') +
        (params.requestId ? `\n  request-id: ${params.requestId}` : ''),
      1,
    );
    this.name = 'GraphError';
    this.status = params.status;
    this.code = params.code;
    this.graphMessage = params.message;
    this.requestId = params.requestId;
    this.retryAfter = params.retryAfter;
  }
}

/**
 * Operational translation of the Graph error codes a mail CLI actually hits
 * (docs/00 §8). Keyed on error.code — never on error.message, which is
 * human-facing prose Microsoft reserves the right to change.
 */
export const GRAPH_ERROR_HINTS: Record<string, string> = {
  InvalidAuthenticationToken:
    'Access token expired or malformed. Run: outlook auth login',
  ErrorAccessDenied:
    'The signed-in account lacks permission. Confirm the Mail.ReadWrite / Mail.Send scopes were consented at login.',
  ErrorItemNotFound:
    'No message with that id. Note: message ids change when a message moves between folders unless immutable ids are used.',
  ErrorInvalidUser: 'The mailbox does not exist or is not licensed for mail.',
  ErrorInvalidRecipients:
    'One or more recipient addresses were rejected. Check for typos and that each is a routable address.',
  ErrorMessageSizeExceeded:
    'Message plus attachments exceeded the mailbox send limit. Use fewer or smaller attachments.',
  ErrorAttachmentSizeExceeded:
    'Inline attachments must be under 3 MB. Larger files need an upload session (not supported in v1).',
  ErrorMailboxStoreUnavailable:
    'The mailbox is temporarily unavailable. Retry shortly.',
  ErrorQuotaExceeded: 'Mailbox is over quota.',
  ErrorFolderNotFound:
    'No folder with that name or id. Run: outlook folders list',
  ErrorInvalidIdMalformed:
    'That id is not a valid Graph id. Ids are opaque — copy them verbatim from a list/get.',
  InefficientFilter:
    'The $filter/$orderby combination is too complex. Properties in $orderby must also appear in $filter, in the same order.',
  activityLimitReached:
    'Throttled by Outlook (429). The client backs off automatically; if this persists, reduce request volume.',
  TooManyRequests: 'Throttled by Outlook (429). Back off and retry.',
  ErrorTooManyObjectsOpened:
    'Too many concurrent mailbox operations. Reduce concurrency to 4 or fewer.',
  ErrorSendAsDenied:
    'The account may not send as that address. Check Send As permissions.',
  MailboxNotEnabledForRESTAPI:
    'This mailbox cannot be accessed via the REST API (often an on-prem or unlicensed mailbox).',
};

/** Parse a Graph error envelope defensively — both innerError casings occur. */
export function parseGraphError(
  status: number,
  bodyText: string,
  retryAfter?: number,
): GraphError {
  let code = `HTTP${status}`;
  let message = bodyText.slice(0, 500) || 'no response body';
  let requestId: string | undefined;

  try {
    const parsed: unknown = JSON.parse(bodyText);
    const root = parsed as {
      error?: unknown;
      error_description?: unknown;
    };

    // Identity-platform errors use flat {error, error_description} strings;
    // Graph errors nest an object under `error`. Handle both.
    if (typeof root.error === 'string') {
      code = root.error;
      if (typeof root.error_description === 'string') {
        message = root.error_description;
      }
    } else if (root.error && typeof root.error === 'object') {
      const err = root.error as {
        code?: unknown;
        message?: unknown;
        innerError?: Record<string, unknown>;
        innererror?: Record<string, unknown>;
      };
      if (typeof err.code === 'string') code = err.code;
      if (typeof err.message === 'string') message = err.message;
      const inner = err.innerError ?? err.innererror;
      const rid = inner?.['request-id'] ?? inner?.['client-request-id'];
      if (typeof rid === 'string') requestId = rid;
    }
  } catch {
    // Non-JSON body (HTML error page, empty 502). Keep the raw slice.
  }

  return new GraphError({ status, code, message, requestId, retryAfter });
}
