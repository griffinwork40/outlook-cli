# Microsoft Graph v1.0 — Authentication & Fundamentals Reference

**Scope:** OAuth2 device code flow, token refresh, tenant/scope semantics, public-client app registration, Outlook-specific throttling, and the Graph error envelope — everything a Node.js CLI needs to authenticate a user and call Outlook mail endpoints.

**Sources (fetched directly, Aug 2026):**
- https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code
- https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow (refresh-token section)
- https://learn.microsoft.com/en-us/graph/permissions-reference
- https://learn.microsoft.com/en-us/graph/throttling-limits
- https://learn.microsoft.com/en-us/graph/throttling
- https://learn.microsoft.com/en-us/graph/errors
- https://www.rfc-editor.org/rfc/rfc8628 (device authorization grant standard — MS implements this)
- https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/app-integration/confidential-client-application-authentication-error-aadsts7000218
- https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc
- https://learn.microsoft.com/en-us/entra/msal/msal-authentication-flows
- https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-application-configuration
- https://learn.microsoft.com/en-us/graph/api/user-sendmail, /message-send, /user-list-messages (mail endpoint examples)
- Microsoft Q&A / GitHub issues (Microsoft-maintained SDK source) used only to corroborate real-world response shapes where the prose docs are silent — flagged inline wherever used.

⚠️ **Read the "Discrepancies & unverifiable items" section (bottom) before implementing polling error handling — the official docs table is incomplete relative to the RFC Microsoft says it implements.**

---

## Table of contents

1. [Tenant values](#1-tenant-values)
2. [Device code flow — device authorization request](#2-device-code-flow--device-authorization-request)
3. [Token polling — the `/token` endpoint](#3-token-polling--the-token-endpoint)
4. [Refresh token exchange](#4-refresh-token-exchange)
5. [Scopes](#5-scopes)
6. [Public client requirement](#6-public-client-requirement)
7. [Throttling (Outlook-specific)](#7-throttling-outlook-specific)
8. [Error envelope + mail error code table](#8-error-envelope--mail-error-code-table)
9. [Discrepancies & unverifiable items](#9-discrepancies--unverifiable-items)
10. [Implementation checklist for the CLI](#10-implementation-checklist-for-the-cli)

---

## 1. Tenant values

The `{tenant}` path segment in every URL below (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/...`) accepts one of these literal values, per the auth-code-flow doc's parameter table (identical text repeated on the device-code page):

| Value | Meaning | Personal MSA (outlook.com/hotmail/live) | Work/school (Entra ID org) accounts |
|---|---|---|---|
| `common` | Sign in users with **either** work/school accounts or personal Microsoft accounts | ✅ | ✅ |
| `organizations` | Sign in users with work/school accounts from **any** Entra ID organization (multi-tenant) | ❌ | ✅ |
| `consumers` | Sign in users with **personal Microsoft accounts only** | ✅ | ❌ |
| `{GUID}` or domain name | Restrict sign-in to **one specific tenant** (single-tenant app) | Only if that tenant permits MSA guests (rare) | ✅ (that tenant only) |

**→ `common` is the only value that supports BOTH personal Microsoft accounts and work/school accounts.** This is the correct choice for a general-purpose CLI tool whose users might have either an `@outlook.com`/`@hotmail.com` account or a corporate/school Microsoft 365 account.

Source: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow (parameter table for `tenant`) and https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-application-configuration ("Common authority URLs" table — identical semantics restated for MSAL).

Caveat from the device-code doc specifically: "If the user authenticates with a personal account, using `/common` or `/consumers`, they're asked to sign in again in order to transfer authentication state to the device… However, this doesn't apply to work or school accounts used to authenticate." — i.e. expect one extra interactive sign-in step for MSA users in the device-code flow specifically.

---

## 2. Device code flow — device authorization request

### Exact endpoint (POST)

```
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode
```

### Request

Content-Type: `application/x-www-form-urlencoded`

| Param | Required | Description |
|---|---|---|
| `client_id` | **Required** | The Application (client) ID from app registration (GUID) |
| `scope` | **Required** | Space-separated scope list, URL-encoded (spaces → `%20` or `+`) |

curl example:

```bash
curl -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "client_id=00001111-aaaa-2222-bbbb-3333cccc4444" \
  --data-urlencode "scope=offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite"
```

### Response (HTTP 200, JSON)

Exact field set per the official parameter table:

| Field | Type | Description |
|---|---|---|
| `device_code` | string | Long opaque string; used by the client to poll `/token`. Never show to the user. |
| `user_code` | string | Short string the **user** types at the verification URL. |
| `verification_uri` | URI | The page the user visits to enter `user_code`. Commonly `https://microsoft.com/devicelogin` in practice (not a documented stable contract — always read the field from the response, never hardcode). |
| `expires_in` | int | Seconds until `device_code`/`user_code` expire. **Default 900 (15 minutes)** per doc prose: "the user has 15 minutes to sign in. This is the default value for `expires_in`." |
| `interval` | int | Minimum seconds the client must wait between polls. Observed value in practice: `5`. |
| `message` | string | Human-readable instructions. Localizable via a `?mkt=xx-XX` query param on the devicecode request. |

Example response body (real shape, ID values redacted/shortened):

```json
{
  "device_code": "GMMhmHCXhWEzkobqIHGG_EnNYYsAkukHspeYUk9E8...",
  "user_code": "FDXXXPQR9",
  "verification_uri": "https://microsoft.com/devicelogin",
  "expires_in": 900,
  "interval": 5,
  "message": "To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code FDXXXPQR9 to authenticate."
}
```

⚠️ **`verification_uri_complete` is explicitly NOT supported by Microsoft**, despite being an optional field in RFC 8628. The docs call this out directly: "The `verification_uri_complete` response field is not included or supported at this time." Do not build a QR-code/deep-link flow expecting it — you must display `user_code` and `verification_uri` separately.

---

## 3. Token polling — the `/token` endpoint

### Exact endpoint (POST)

```
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
```

### Exact `grant_type` value — **this is the long URN form, copy it exactly:**

```
urn:ietf:params:oauth:grant-type:device_code
```

(Verified byte-for-byte identical across the device-code doc's prose example, the auth-code-flow doc's cross-reference, and RFC 8628 §3.4's own registered grant-type URN. There is no shorthand — `device_code` alone as a `grant_type` value will fail.)

### Request params

| Param | Required | Description |
|---|---|---|
| `grant_type` | **Required** | Literal string `urn:ietf:params:oauth:grant-type:device_code` |
| `client_id` | **Required** | Must match the `client_id` used in the devicecode request |
| `device_code` | **Required** | The `device_code` value returned by step 2 |

curl example:

```bash
curl -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "grant_type=urn:ietf:params:oauth:grant-type:device_code" \
  --data-urlencode "client_id=00001111-aaaa-2222-bbbb-3333cccc4444" \
  --data-urlencode "device_code=GMMhmHCXhWEzkobqIHGG_EnNYYsAkukHspeYUk9E8..."
```

### Polling error codes — client action table

While the user hasn't finished authenticating, the token endpoint returns **HTTP 400** with a JSON error body `{"error": "<code>", "error_description": "...", ...}`. Per the official doc's table (device-code page) plus RFC 8628 §3.5 (see the discrepancy flag below for `slow_down`):

| `error` value | Meaning | Client action |
|---|---|---|
| `authorization_pending` | User hasn't finished authenticating, hasn't cancelled | **Repeat the request after at least `interval` seconds.** Keep the same `device_code`. |
| `slow_down` | ⚠️ See discrepancy note below — RFC 8628-defined, NOT in MS's own docs table | **Increase your polling interval by 5 seconds** (permanently, for this and all subsequent polls) and continue polling. Do not treat as fatal. |
| `expired_token` | `expires_in` window elapsed; `device_code` is dead | **Stop polling.** Revert to unauthenticated state. Start a new devicecode request only when the user is ready again. |
| `authorization_declined` | User explicitly denied the consent/sign-in prompt | **Stop polling.** Revert to unauthenticated state. Do not auto-retry. |
| `bad_verification_code` | The `device_code` sent wasn't recognized by the server | Verify you're sending the exact `device_code` string from step 2 (no truncation/whitespace). This should not occur in normal operation — treat as a bug in your polling loop, not a transient condition. |

Example pending response:

```json
{
  "error": "authorization_pending",
  "error_description": "AADSTS70016: The provided request has not yet been authorized by the user. The user must input their code.\r\nTrace ID: ...\r\nCorrelation ID: ...\r\nTimestamp: ...",
  "error_codes": [70016],
  "timestamp": "2022-08-25 09:34:07Z",
  "trace_id": "...",
  "correlation_id": "..."
}
```

Only the `error` field's string value is stable/contractual. `error_description` (with its embedded `AADSTS#####` code) is for human-readable diagnostics — never branch your logic on its text.

### Successful token response (HTTP 200)

```json
{
  "token_type": "Bearer",
  "scope": "User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite profile openid email",
  "expires_in": 3599,
  "access_token": "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik5HVEZ2ZEstZnl0aEV1Q...",
  "refresh_token": "AwABAAAAvPM1KaPlrEqdFSBzjqfTGAMxZGUTdM0t4B4...",
  "id_token": "eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIyZDRkMTFhMi1mODE0LTQ2YTctOD..."
}
```

| Field | Description |
|---|---|
| `token_type` | Always `Bearer` |
| `scope` | Space-separated scopes the `access_token` is actually valid for |
| `expires_in` | Access token lifetime in seconds (observed default ~3599 = 1 hour) |
| `access_token` | Opaque/JWT-like string. **Do not parse or take a dependency on its internal structure** — official warning: "Tokens for Microsoft services can use a special format that will not validate as a JWT, and may also be encrypted for consumer (Microsoft account) users." |
| `refresh_token` | Present **only if `offline_access` was requested** in scope |
| `id_token` | Present **only if `openid` was requested** in scope |

Note: some real-world responses also include an `ext_expires_in` field — this is **not** in the official parameter table for either the device-code or auth-code-flow docs; treat it as optional/undocumented (see §9).

---

## 4. Refresh token exchange

### Endpoint

Same `/token` endpoint as above:

```
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
```

### Exact params for `grant_type=refresh_token`

| Param | Required | Description |
|---|---|---|
| `grant_type` | **Required** | Literal string `refresh_token` |
| `client_id` | **Required** | Same client ID |
| `refresh_token` | **Required** | The refresh token from the previous token response |
| `scope` | optional | Space-separated list; must be equal to or a subset of the originally granted scopes |

For a **public client** (device code flow), no `client_secret` is sent (see §6).

curl example:

```bash
curl -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "client_id=00001111-aaaa-2222-bbbb-3333cccc4444" \
  --data-urlencode "grant_type=refresh_token" \
  --data-urlencode "refresh_token=AwABAAAAvPM1KaPlrEqdFSBzjqfTGAMxZGUTdM0t4B4..." \
  --data-urlencode "scope=offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite"
```

Response shape is identical to the initial token response (§3) — `access_token`, `token_type`, `expires_in`, `scope`, `refresh_token` (if `offline_access` requested), `id_token` (if `openid` requested).

### Does Microsoft rotate the refresh token? **Yes — implement rotation. Always persist the newest one.**

Direct quote from the official doc: *"Refresh tokens aren't revoked when used to acquire new access tokens. You're expected to discard the old refresh token."* It then cites RFC 6749 §6: *"The authorization server MAY issue a new refresh token, in which case the client MUST discard the old refresh token and replace it with the new refresh token. The authorization server MAY revoke the old refresh token after issuing a new refresh token to the client."*

The field-level description in the doc is stronger/more concrete than the "MAY" language suggests in practice: *"refresh_token: A new OAuth 2.0 refresh token. **Replace the old refresh token with this newly acquired refresh token to ensure your refresh tokens remain valid for as long as possible.**"*

**Implementation implication:** every time your CLI calls `grant_type=refresh_token`, overwrite your stored `refresh_token` with whatever comes back in the response — never keep reusing the original one. Refresh tokens are long-lived (typically ~90 days for confidential/public clients; no fixed lifetime is officially guaranteed) but can be silently invalidated if you keep using a stale one after rotation.

---

## 5. Scopes

### Exact scope strings requested

```
offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite
```

### Do they need the full `https://graph.microsoft.com/` prefix?

**No — not strictly required, but recommended for explicitness.** Per the official scopes doc: *"In requests to the authorization, token or consent endpoints for the Microsoft identity platform, if the resource identifier is omitted in the scope parameter, the resource is assumed to be Microsoft Graph. For example, `scope=User.Read` is equivalent to `https://graph.microsoft.com/User.Read`."*

So both of these are equivalent for Graph scopes:
```
scope=offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite
scope=offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/MailboxSettings.ReadWrite
```

`offline_access` (and `openid`/`profile`/`email`) are OIDC scopes and **never** take a resource prefix — they are resource-agnostic by definition.

Exception: if you ever combine a Graph scope with a **non-Graph** resource's scope in the same request (e.g. also requesting an Azure Key Vault scope), you **must** use the fully-prefixed form for disambiguation, since the shorthand only works when Graph is the sole/implied resource.

### What each scope grants

| Scope | Type | AdminConsentRequired (delegated) | Grants |
|---|---|---|---|
| `offline_access` | OIDC | No | Returns a `refresh_token` in the response so the app can silently get new access tokens without re-prompting the user. Not a Graph data permission by itself. |
| `User.Read` | Delegated | No | *"Allows users to sign-in to the app, and allows the app to read the profile of signed-in users. It also allows the app to read basic company information of signed-in users."* Also unlocks the `organization` resource's `id`/`displayName`/`verifiedDomains` for work/school accounts. |
| `Mail.ReadWrite` | Delegated | No | *"Allows the app to create, read, update, and delete email in user mailboxes. Does not include permission to send mail."* |
| `Mail.Send` | Delegated | No | *"Allows the app to send mail as users in the organization."* Per docs: *"With the `Mail.Send` permission, an app can send mail and save a copy to the user's Sent Items folder, even if the app isn't granted the `Mail.ReadWrite`… permission."* — i.e. `Mail.Send` alone is sufficient for `sendMail`/`send`; you don't strictly need `Mail.ReadWrite` too, but you do need it for listing/reading/updating drafts. |
| `MailboxSettings.ReadWrite` | Delegated | No | *"Allows the app to read and update user's mailbox settings."* Needed for `rules create`; `MailboxSettings.Read` suffices for read-only operations. |

All five scopes above are confirmed **available for consent in personal Microsoft accounts** (each carries the MSA-eligible badge in the permissions reference) and confirmed valid for both work/school and personal accounts — consistent with using tenant `common`.

Source: https://learn.microsoft.com/en-us/graph/permissions-reference (sections: Mail.Read, Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite, User.Read — each fetched and cross-checked directly).

---

## 6. Public client requirement

### What "Allow public client flows" does

In the app registration (**Authentication** blade → **Advanced settings** → **Allow public client flows**), this setting tells Microsoft Entra ID that the app is a **public client** — one that cannot securely store a secret (native/CLI/mobile apps, as opposed to server-side web apps that can hold a `client_secret`).

This setting (or equivalently, registering the app with a "Mobile and desktop applications" platform / no Web redirect URI) is **specifically required for**:
- OAuth2 Device Code flow (this doc's flow)
- Resource Owner Password Credentials (ROPC) flow
- Integrated Windows Authentication

Per Microsoft's own MSAL docs: *"The device code flow is available only for public client applications."*

Mechanically, Entra ID determines public-vs-confidential client status from either (a) the type of redirect URI registered (Web = confidential, "Mobile and desktop applications" = public), or (b) for flows with no redirect URI at all (device code, ROPC), from the **"Allow public client flows"** toggle directly.

### What error you get if it's off

**`AADSTS7000218`** — HTTP 400, `"error": "invalid_client"`:

```json
{
  "error": "invalid_client",
  "error_description": "AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'.\r\nTrace ID: ...\r\nCorrelation ID: ...\r\nTimestamp: ...",
  "error_codes": [7000218],
  "timestamp": "...",
  "trace_id": "...",
  "correlation_id": "..."
}
```

This occurs because, with the toggle off (or default), Entra ID treats the app registration as a **confidential client**, and confidential clients must always authenticate to the token endpoint with a `client_secret` or `client_assertion` (certificate/JWT) — which a device-code-flow CLI, by design, does not and should not send. Turning on "Allow public client flows" (Yes) removes this requirement for device-code and ROPC requests specifically.

Fix: App registration → **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes** → Save.

Source: https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/app-integration/confidential-client-application-authentication-error-aadsts7000218 (official MS troubleshooting article, fetched directly) plus corroborating community threads (Microsoft Q&A, GitHub AzureAD/microsoft-authentication-library issues) all converging on the identical fix.

---

## 7. Throttling (Outlook-specific)

### Global ceiling (applies above everything else)

Per-app, across all tenants: **130,000 requests per 10 seconds** (any request type). This is the outermost limit; Outlook-specific limits below bind tighter and will trigger first in normal CLI usage.

### Outlook mailbox-scoped limits

Per the official throttling-limits doc, the Outlook service throttles **per (app ID, mailbox) pair** — i.e. hitting the limit against one user's mailbox does not affect the app's ability to call a different user's mailbox:

| Limit | Applies to |
|---|---|
| **10,000 API requests per 10-minute period** | v1.0 and beta endpoints |
| **4 concurrent requests** | v1.0 and beta endpoints |
| **150 MB upload (PATCH, POST, PUT) per 5-minute period** | v1.0 and beta endpoints |

These limits cover the resources your CLI will touch: `message`, `mailFolder`, `mailSearchFolder`, `messageRule`, `outlookCategory`, `attachment` (Mail API group) — plus `event`/`calendar` if you ever add calendar support.

### JSON batch behavior against the same mailbox

If you batch multiple unordered requests via `POST /$batch`, Microsoft Graph forwards **at most 4 individual requests at a time** to the Outlook backend — deliberately matching the 4-concurrent-request mailbox limit above, regardless of how many requests are in your batch. If you need requests to execute in a specific order, use the batch's `dependsOn` property; Graph then sends them to Outlook **one at a time**, sequentially.

### The 429 response shape

Exact JSON body from the official throttling doc:

```
HTTP/1.1 429 Too Many Requests
Content-Length: 312
Content-Type: application/json
Retry-After: 10

{
  "error": {
    "code": "TooManyRequests",
    "innerError": {
      "code": "429",
      "date": "2020-08-18T12:51:51",
      "message": "Please retry after",
      "request-id": "94fb3b52-452a-4535-a601-69e0a90e3aa2",
      "status": "429"
    },
    "message": "Please retry again later."
  }
}
```

### `Retry-After` header semantics

- Value is in **seconds**.
- Per the docs: *"Backing off requests using the `Retry-After` delay is the fastest way to recover from throttling because Microsoft Graph continues to log resource usage while a client is being throttled."* i.e. retrying before the delay elapses actively extends your throttled window.
- **All Outlook Mail API resources return a `Retry-After` header** on 429 (the doc explicitly calls out only a small set of *other* Graph resources — reports, dataPolicyOperation, risk/conditional-access resources, OneNote — as the exceptions that do NOT return `Retry-After`; Outlook mail is not in that exception list).

### Recommended backoff

1. Wait exactly the number of seconds in `Retry-After`.
2. Retry the request once.
3. If still 429, you're still throttled — **keep honoring the new `Retry-After` value** on each subsequent 429; don't apply your own multiplier on top of it.
4. Only if `Retry-After` is absent (shouldn't happen for Outlook mail endpoints per above, but defend anyway) fall back to **exponential backoff with jitter**, per the doc's general guidance for the no-header case.
5. Never batch-retry a whole failed JSON-batch request — retry only the individual sub-requests that came back 429, each waiting its own `Retry-After` value (per-request retry-after inside the batch response), and re-batch only after the *longest* of those individual delays.

Sources: https://learn.microsoft.com/en-us/graph/throttling (mechanism + backoff policy) and https://learn.microsoft.com/en-us/graph/throttling-limits (Outlook-specific numeric limits, both fetched directly).

---

## 8. Error envelope + mail error code table

### Exact JSON shape

Two things to note carefully — **the official docs page itself is internally inconsistent on casing**, flagged in §9. The worked example on the page uses `innerError` (capital E):

```json
{
  "error": {
    "code": "badRequest",
    "message": "Uploaded fragment overlaps with existing data.",
    "innerError": {
      "code": "invalidRange",
      "request-id": "request-id",
      "date": "date-time"
    }
  }
}
```

...but the formal "error resource" JSON schema on the same page spells it `innererror` (lowercase e):

```json
{
  "error": {
    "code": "string",
    "message": "string",
    "innererror": {
      "code": "string"
    },
    "details": []
  }
}
```

**Implementation guidance:** parse defensively — check for both `innerError` and `innererror` keys (case-insensitively, or check both explicitly) since real API responses have been observed using `innerError` (matching the worked example, and matching what's shown in this doc's device-code/token-endpoint error examples too). Do not hard-fail if the key is absent; it's documented as optional.

| Property | Type | Description |
|---|---|---|
| `error.code` | string | Machine-readable error code. **This is the only field your code should branch on.** |
| `error.message` | string | Human-readable, developer-facing. *"Don't take any dependency on the content of this value in your code… they can change at any time."* Never shown to end users. |
| `error.innerError` (or `innererror`) | object | Optional, more specific nested error. Can itself recurse with further nested `innerError`/`innererror` objects — *"apps should loop through all the nested error codes… and use the most detailed one that they understand."* |
| `error.details` | array | Optional array of same-shaped error objects — populated when a single call represents multiple sub-operations (e.g. batch/bulk), one error entry per failed sub-operation. |

### HTTP status codes (Graph-wide, from the same page)

| Code | Meaning |
|---|---|
| 400 | Bad Request — malformed/incorrect request |
| 401 | Unauthorized — missing/invalid/expired token |
| 403 | Forbidden — insufficient permission, missing license, or **conditional access**: watch for `HTTP 403; Forbidden error=insufficient_claims` specifically, which means a CA policy (e.g. MFA) blocked the token and needs a claims-challenge re-auth, not a scope fix |
| 404 | Not Found |
| 409 | Conflict — e.g. concurrent modification; safe to retry with backoff, honor `Retry-After` if present |
| 413 | Request Entity Too Large |
| 429 | Too Many Requests — see §7 |
| 503 | Service Unavailable — transient; retry after `Retry-After` if present |

### Common mail-related error codes (operational table)

These are Exchange/EWS-heritage `error.code` strings (PascalCase `Error*` or bare PascalCase), distinct in style from the Graph-native camelCase codes (`badRequest`, `TooManyRequests`) — both styles appear in the wild depending on which backend service produced the error. Verified against official EWS/Exchange docs plus corroborating first-party Microsoft Q&A threads:

| `error.code` | HTTP status | Meaning | Operational action |
|---|---|---|---|
| `ErrorItemNotFound` | 404 (sometimes reported as 403 in-body) | *"The specified object was not found in the store, or the store ID provided isn't an ID of an item."* Common causes: item ID is stale (moved/copied since you fetched it — message/folder IDs are not guaranteed stable across moves; use [immutable IDs](https://learn.microsoft.com/en-us/graph/outlook-immutable-id) if you need durability), or you're calling `/reply`, `/forward`, `/send` against an ID from the wrong mailbox context. | Refetch the item's current ID before retrying; do not blindly retry with the same ID. |
| `ErrorAccessDenied` | 403 | *"Access is denied. Check credentials and try again."* Causes: missing/unconsented scope, EWS organizational access policy (`EwsApplicationAccessPolicy` allow/block list) blocking your app's User-Agent, or attempting cross-mailbox access without the right delegated/application permission and (for shared mailboxes) missing the `X-AnchorMailbox` header. | Verify granted scopes via the access token (do NOT decode/validate it programmatically per the tokens warning above, but a debug-only manual jwt.io check is fine); confirm admin consent was actually granted, not just requested. |
| `ErrorInvalidRecipients` | 400/403 (docs show it returned in the `code` field regardless of a 200-looking wrapper in some client libraries) | *"At least one recipient is not valid"* / *"the recipient collection… is invalid."* Real-world cause seen most often: `toRecipients` sent as a bare object instead of an array, or calling the wrong send endpoint for a draft (`sendMail` action expects a full new message; an existing draft must use `POST /messages/{id}/send` with an empty body, not `sendMail`). | Validate your JSON matches the documented shape exactly: `toRecipients: [{ "emailAddress": { "address": "..." } }]` (array, not object). If sending an existing draft, switch to `/messages/{id}/send`. |
| `ErrorMessageSizeExceeded` | 403/413 | Attachment or message body exceeds the applicable size ceiling. Individual `POST` of a file attachment via JSON is capped at ~4 MB in practice / the Graph large-attachment doc's threshold is **3 MB** as the cutover point to an upload session; total mailbox send/receive size defaults to **35 MB send / 36 MB receive** (org-configurable up to 150 MB). | For attachments ≥3 MB, use the [upload session](https://learn.microsoft.com/en-us/graph/outlook-large-attachments) (`createUploadSession` + chunked `PUT`) instead of a single POST. |
| `ErrorMimeContentInvalidBase64String` | 400 | MIME-format `sendMail`/`send` body wasn't valid base64. | Only relevant if you support raw-MIME send; verify base64 encoding before sending. |
| `InvalidAuthenticationToken` | 401 | Graph-native (camelCase-family) code, message commonly `"CompactToken parsing failed with error code: <hex>"`. Causes: missing `Bearer` prefix, malformed/truncated token, token issued for the wrong audience/resource (e.g. an ID token or a token for a different API used where a Graph access token is expected), or an expired token being reused past its `expires_in`. | Confirm the `Authorization` header is exactly `Bearer {access_token}` (one space, no extra whitespace/newlines); confirm the token's intended audience is `https://graph.microsoft.com` (decode for debugging only); refresh and retry once — if it recurs immediately after a fresh token, it's a code/audience bug, not an expiry issue. |
| `ErrorQuotaExceeded` | 403 | Ambiguous in practice — message varies: *"Mailbox has exceeded maximum mailbox size"*, *"Cannot get ID from name"* (seen specifically when using `singleValueExtendedProperties` with a malformed/over-quota property ID), or per-folder item-count ceiling exceeded (delta query on a folder with excessive items). | Don't assume literal storage quota — read `error.message` for the specific sub-cause; if using `singleValueExtendedProperties`, try removing it first to isolate; check folder item counts against [service-specific limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits) if delta query is involved. |
| `MailboxNotEnabledForRESTAPI` / `MailboxNotSupportedForRESTAPI` | 404 | Target mailbox isn't a valid Exchange Online / Graph-supported mailbox (e.g. a mailbox still on a dedicated/legacy Exchange Server not yet migrated). | Not fixable client-side; requires the mailbox to be migrated to Exchange Online by an admin. Surface as a hard failure, not a retry candidate. |
| `ErrorInvalidMailboxItemId` | 404 | Item ID belongs to a different mailbox than the one addressed in the URL (cross-mailbox ID reuse across users is blocked). Also returned for **in-place archive folders** (`ArchiveMsgFolderRoot`, `ArchiveDeletedItems`, etc.) — the Mail API [explicitly does not support archive mailboxes](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview). | Confirm the mailbox segment in the URL (`/me/...` vs `/users/{id}/...`) matches the mailbox that actually owns the item; don't attempt archive-folder access. |
| `MailboxInfoStale` / plain `503 Service Unavailable` | 503 | Transient backend routing issue (mailbox database/backend temporarily "Unknown" to the front-end). | Retry with backoff honoring `Retry-After` if present (per §7); typically self-resolves. |
| `ErrorAttachmentSizeShouldNotBeLessThanMinimumSize` | 400 | You called `createUploadSession` for a file under 3 MB. | Use the direct single-`POST` attachment path instead of an upload session for files <3 MB. |
| `NoPermissionsInAccessToken` | 403 | Token has no recognized/parseable `scp` (delegated scopes) claim for the resource being called — commonly a tenant/endpoint mismatch (e.g. authenticated against `/organizations` but the account is an MSA, or vice-versa) even though the requested scopes looked correct in the authorize step. | Verify the `tenant` segment used for BOTH `/devicecode` and `/token` matches your intended audience (`common`, per §1) and that the decoded token's `scp` claim actually lists the expected scopes — if empty, the consent/grant didn't apply to the resource you're calling. |

Sources: https://learn.microsoft.com/en-us/graph/errors (envelope schema + HTTP codes, fetched directly), https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxwscdata (ResponseCodeType enum — official protocol spec for the `Error*` code family's canonical meanings), https://learn.microsoft.com/en-us/graph/outlook-large-attachments, https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview, plus Microsoft Q&A threads used only to corroborate the *practical* trigger conditions the prose docs don't spell out (each flagged inline above where used).

---

## 9. Discrepancies & unverifiable items

Flagging these explicitly per the research brief — do not treat any of the following as officially nailed-down without the caveats noted:

1. **`slow_down` is NOT in Microsoft's own device-code-flow docs table, despite Microsoft saying it implements RFC 8628 (which defines `slow_down`).** I fetched https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code twice directly (byte-identical both times) — its "Expected errors" table lists exactly four codes: `authorization_pending`, `authorization_declined`, `bad_verification_code`, `expired_token`. `slow_down` appears nowhere on that page. However:
   - RFC 8628 §3.5 (the standard the page cites and says Microsoft implements) explicitly defines `slow_down` as a required-to-handle variant of `authorization_pending`.
   - Microsoft's own Azure SDK source (`Azure/go-autorest`, an official Microsoft GitHub org repo, `autorest/adal/devicetoken.go`) contains explicit handling for `"slow_down"` returned by Azure AD's device-code token endpoint, with logic to increase backoff.
   - This is circumstantial-but-strong evidence Microsoft's endpoint *can* send `slow_down` in practice even though the current customer-facing doc omits it from the table.
   - **Recommendation implemented in this doc (§3):** handle `slow_down` defensively (treat like `authorization_pending` but permanently add 5s to your polling interval) even though it's undocumented on the current page — cost of handling it is near-zero, cost of NOT handling it (treating an unrecognized error as fatal) is a false auth failure.

2. **`ext_expires_in` on the token response** — appears in real-world example payloads (Microsoft Q&A threads) but is absent from the official parameter tables on both the device-code page and the auth-code-flow page for either the initial-grant or refresh-token responses. Treat as optional/non-contractual; do not require it.

3. **`innerError` vs `innererror` casing** — the official `/graph/errors` page itself uses `innerError` (capital E) in its worked JSON example but `innererror` (lowercase) in its formal schema definition immediately below it, on the same page. This is a genuine inconsistency in Microsoft's own documentation, not a research gap on my part — parse defensively for both.

4. **`verification_uri`'s concrete value** (`https://microsoft.com/devicelogin`) is only ever shown in examples/community answers, never asserted as a stable contract in the parameter table itself (which just says "The URI the user should go to"). Treat it as **read from the response every time**, never hardcoded, even though it has been observed unchanged for years.

5. **Exact default numeric value of `expires_in` for the devicecode response** — the doc states it in prose ("15 minutes") but does not print a literal `900` in the parameter table itself; `900` is corroborated only via multiple real observed response bodies (Microsoft Q&A). Treat 900s as the practical default, not a documented numeric guarantee that could never change.

6. **`ErrorQuotaExceeded`'s "Cannot get ID from name" message** for the unrelated `singleValueExtendedProperties` case is corroborated only by Microsoft Q&A/GitHub issue threads (developer-reported), not by an official reference page enumerating this specific mapping. Flagged as community-sourced in the table in §8.

7. **Could not find an official, single canonical "Graph mail error code" reference page** enumerating `ErrorItemNotFound`/`ErrorAccessDenied`/`ErrorInvalidRecipients`/etc. as a first-class Graph-documented list (unlike the generic `/graph/errors` envelope page, which is Graph-native and doesn't enumerate Exchange-heritage codes at all). The table in §8 is synthesized from the official EWS `ResponseCodeType` protocol spec (which does canonically define these strings and their meanings) cross-referenced against official Graph how-to articles (large-attachments, mail-api-overview) and corroborating Q&A for practical trigger conditions. This is the best available official grounding, but it is not one single MS Graph page — flagging so the implementer knows to expect codes not in this table if new Exchange-backend errors surface.

---

## 10. Implementation checklist for the CLI

1. **App registration (one-time, human step):** register an app, set **Supported account types → Accounts in any organizational directory and personal Microsoft accounts**, enable **Allow public client flows = Yes** (§6), grant delegated scopes `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite` (no admin consent needed per §5 — all are user-consentable). No redirect URI is needed for device-code flow.
2. **Auth module:**
   - `POST /devicecode` with `tenant=common` (§1, §2) → display `message`/`user_code`/`verification_uri` to the user.
   - Poll `POST /token` with `grant_type=urn:ietf:params:oauth:grant-type:device_code` (§3) every `interval` seconds; handle all 5 polling codes including the undocumented-but-real `slow_down` (§9.1).
   - Persist `access_token`, `refresh_token`, and computed expiry (`now + expires_in`) to local encrypted storage (OS keychain preferred for a Node CLI — e.g. `keytar`).
3. **Token refresh:** before each API call, if the access token is expired (or proactively a minute early), call `grant_type=refresh_token` (§4) and **overwrite the stored refresh token** with whatever comes back — never keep reusing the original.
4. **API calls:** `Authorization: Bearer {access_token}` header; base URL `https://graph.microsoft.com/v1.0/`.
5. **Throttling:** wrap outbound calls with a `Retry-After`-aware retry (§7); cap concurrency at ≤4 simultaneous requests per mailbox to stay under the Outlook concurrency ceiling proactively rather than reactively.
6. **Error handling:** parse `error.code` (never `error.message`) for branching logic (§8); check both `innerError`/`innererror` casings (§9.3); map the mail-specific codes in §8's table to user-facing CLI messages.
