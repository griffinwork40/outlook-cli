# Setup — one-time Azure app registration

`outlook-cli` talks to your mailbox through the Microsoft Graph API. Microsoft
requires every Graph client to be a registered application, so there is exactly
one step a human has to do that no script can do for you: **register an app and
copy its client ID.**

It takes about three minutes, costs nothing, and needs no Azure subscription —
a plain Microsoft account (outlook.com, hotmail, live) or a work/school account
is enough.

---

## Why this step exists

There is no "just give me an API key" path for Outlook mail. Microsoft issues
credentials to *applications*, not to users, and a token is only ever minted
after you personally sign in and consent. That is also what keeps this safe:
the CLI never sees your password, and you can revoke its access at any time
without changing your password.

The client ID you create is **not a secret**. It identifies the app, not you.
The device-code flow this CLI uses is designed for exactly this case — a
command-line tool on a machine you control, with no client secret to leak.

---

## Step 1 — Open the app registration page

Go to <https://entra.microsoft.com> → **Applications** → **App registrations**.

Direct link: <https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade>

(The older <https://portal.azure.com> → *Azure Active Directory* → *App
registrations* path leads to the same screen.)

Sign in with the Microsoft account whose mail you want the CLI to read.

## Step 2 — New registration

Click **+ New registration**.

- **Name:** anything — `outlook-cli` is fine. Only you ever see it.
- **Supported account types:** choose

  > **Accounts in any organizational directory (Any Microsoft Entra ID tenant —
  > Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**

  This is the option that makes one registration work for BOTH a personal
  outlook.com mailbox and a work/school Microsoft 365 mailbox. Picking a
  narrower option is the single most common reason sign-in later fails with
  *"unauthorized_client"*.
- **Redirect URI:** leave it **blank**. The device-code flow does not use one.

Click **Register**.

## Step 3 — Copy the client ID

On the app's **Overview** page, copy **Application (client) ID**. It looks like:

```
11111111-2222-3333-4444-555555555555
```

That is your `GRAPH_CLIENT_ID`.

## Step 4 — Turn on public client flows

This is the step people miss, and it is required.

Go to **Authentication** (left sidebar) → scroll to **Advanced settings** →
**Allow public client flows** → set to **Yes** → **Save**.

Without this, `outlook auth login` fails with:

```
AADSTS7000218: The request body must contain the following parameter:
'client_assertion' or 'client_secret'.
```

That error means precisely this toggle is off — not that you need a secret.

## Step 5 — Export the client ID and sign in

```bash
export GRAPH_CLIENT_ID=11111111-2222-3333-4444-555555555555
outlook auth login
```

You will see something like:

```
To sign in, open https://microsoft.com/devicelogin and enter code ABCD-EFGH
```

Open that URL, enter the code, sign in, and approve the permission list. The
CLI polls in the background and confirms once you are done. The token is cached
at `~/.afk/state/outlook/token.json` (mode `0600`) and refreshes itself from
then on — you should not need to log in again unless you revoke access.

Verify:

```bash
outlook auth status
outlook me
```

---

## Making it permanent

Add the export to your shell profile, or drop it in AFK's env file so every
session picks it up:

```bash
echo 'GRAPH_CLIENT_ID=11111111-2222-3333-4444-555555555555' >> ~/.afk/config/afk.env
```

Optional variables (see `.env.example`):

| Variable | Default | Purpose |
|---|---|---|
| `GRAPH_CLIENT_ID` | *(required)* | Application (client) ID from Step 3 |
| `GRAPH_TENANT_ID` | `common` | `common` supports personal **and** work accounts |
| `GRAPH_CLIENT_SECRET` | *(unset)* | Only for confidential-client mode; leave blank |
| `GRAPH_TOKEN_CACHE` | `~/.afk/state/outlook/token.json` | Token cache location |
| `OUTLOOK_SIGNATURE_FILE` | *(unset)* | Plain-text signature appended to outgoing mail (see README "Signatures") |
| `OUTLOOK_SIGNATURE_HTML_FILE` | *(unset)* | HTML signature used for `--html` bodies |

---

## What the CLI asks permission for

At first sign-in Microsoft shows a consent screen listing these scopes:

| Scope | Why it is needed |
|---|---|
| `offline_access` | Refresh the token so you sign in once, not daily |
| `User.Read` | Read your own profile — powers `outlook me` |
| `Mail.ReadWrite` | Read messages, manage drafts, flags, folders |
| `Mail.Send` | Send, reply, forward |
| `MailboxSettings.ReadWrite` | Read/write mailbox settings (needed for `rules create`) |

These are **delegated** permissions: the CLI can only ever do what *you* can do,
and only while your token is valid.

## Revoking access

<https://account.live.com/consent/Manage> (personal accounts) or your tenant's
*Enterprise applications* blade (work accounts). Revoking is immediate and does
not require deleting the app registration.

To just clear the local token:

```bash
outlook auth logout
```

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `AADSTS7000218 ... client_assertion or client_secret` | Public client flows off | Step 4 |
| `AADSTS700016 ... application not found` | Wrong or unset `GRAPH_CLIENT_ID` | Re-copy from Step 3 |
| `unauthorized_client` on a personal account | Registration is single-tenant | Re-do Step 2 with the multitenant + personal option |
| `Missing required environment variable(s): GRAPH_CLIENT_ID` | Not exported | Step 5 |
| `InvalidAuthenticationToken` on a call | Token expired and refresh failed | `outlook auth login` |
| Sign-in works, mail calls 403 | Consent was declined or partial | `outlook auth logout` then log in again and accept all scopes |
