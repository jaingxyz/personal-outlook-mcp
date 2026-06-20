# personal-outlook-mcp

[![CI](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/ci.yml)
[![CodeQL](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/codeql.yml/badge.svg)](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/codeql.yml)
[![Semgrep](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/semgrep.yml/badge.svg)](https://github.com/jaingxyz/personal-outlook-mcp/actions/workflows/semgrep.yml)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/jaingxyz/personal-outlook-mcp/badge)](https://scorecard.dev/viewer/?uri=github.com/jaingxyz/personal-outlook-mcp)
[![npm version](https://img.shields.io/npm/v/@jaingxyz/personal-outlook-mcp?label=npm)](https://www.npmjs.com/package/@jaingxyz/personal-outlook-mcp)
[![Smithery](https://smithery.ai/badge/jaingxyz/personal-outlook-mcp)](https://smithery.ai/servers/jaingxyz/personal-outlook-mcp)
[![License: AGPL v3+](https://img.shields.io/badge/license-AGPL--3.0--or--later-blue)](./LICENSE)

A [Model Context Protocol](https://modelcontextprotocol.io/) server that exposes a personal Outlook (consumer Microsoft account) inbox **and calendar** to MCP clients like Claude Desktop. Talks to Microsoft Graph over HTTPS, uses MSAL device-code flow for OAuth, and stores tokens in the OS keyring.

## How it works

Each user runs the server locally on their own machine and connects it to their own personal Microsoft account. There is no hosted version. **You bring your own Azure AD app registration** — see "Azure setup" below. This is intentional: your tokens never leave your machine, throttling is your own, and the consent screen shows an app you own (not someone else's).

> **"Personal" here describes the _account type_, not a no-setup mode.** It means a consumer Microsoft account (`@outlook.com`, `@hotmail.com`, `@live.com`) as opposed to a work/school account. It does **not** mean a shared/hosted client you can use without configuration — you still complete the one-time Azure app registration below to get an `AZURE_CLIENT_ID`. Likewise, the **device-code flow** is just the sign-in _style_; it still authenticates through your app registration, so it can't be used to skip that step.

## Quickstart

```bash
# 1. Register your own Azure AD app (~5 minutes, see "Azure setup" below).
#    Copy the Application (client) ID.

# 2. Add to ~/Library/Application Support/Claude/claude_desktop_config.json:
{
  "mcpServers": {
    "personal-outlook": {
      "command": "npx",
      "args": ["-y", "@jaingxyz/personal-outlook-mcp"],
      "env": {
        "AZURE_CLIENT_ID": "<YOUR-CLIENT-ID>",
        "AZURE_TENANT": "consumers"
      }
    }
  }
}

# 3. From a terminal, seed the OS keyring once:
npx -y @jaingxyz/personal-outlook-mcp whoami
# Paste the device code in the printed URL, sign in, approve.

# 4. Restart Claude Desktop. Tools appear under personal_email_* and personal_calendar_*.
```

## Prerequisites

- Node.js **22.22.1+** (tested against Node 22 and 24).
- An Azure AD app registration's **Application (client) ID** — instructions below.
- One of the supported keyring backends:
  - **macOS** — built-in Keychain. No setup.
  - **Windows** — built-in Credential Manager. No setup.
  - **Linux** — a Secret Service implementation (e.g. `gnome-keyring` or `kwallet`) running.

## Azure setup

You **must** register your own Azure AD app. This server has no hosted backend — every user is their own publisher. Cost: $0, time: ~5 minutes.

> **You do NOT need a paid Azure subscription.** Sign into https://entra.microsoft.com with your personal Microsoft account and a free "Default Directory" tenant is created for you automatically; registering an app in it costs nothing. Microsoft's generic quickstart says an "active subscription" is required — that does **not** apply to registering an app with a personal account.

1. Go to https://entra.microsoft.com (sign in with your personal Microsoft account) → **App registrations** → "New registration". (`portal.azure.com` works too, but the Entra admin center is smoother for personal accounts.)
2. **[REQUIRED]** **Supported account types**: "Accounts in any organizational directory **and** personal Microsoft accounts". This is load-bearing — the `consumers` authority this server uses only accepts personal Microsoft accounts. Pick the wrong option here and sign-in fails instantly with `invalid_grant` and no device code (see Troubleshooting).
3. Leave the redirect URI blank. Click Register.
4. In the new app: **Authentication** → "Add a platform" → "Mobile and desktop applications" → add the redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`.
5. **[REQUIRED]** Same page, scroll to **Advanced settings** → set "Allow public client flows" = **Yes**. Save. Device-code flow is a public-client flow; without this, sign-in fails with `invalid_grant`.
6. Copy the **Application (client) ID** from the Overview page.

The required Graph delegated permissions (`Mail.ReadWrite`, `Mail.Send`, `offline_access`, `User.Read`) are requested at sign-in. For personal accounts the user consents at the device-code prompt — no admin consent required.

> If `portal.azure.com` (or `entra.microsoft.com`) returns `AADSTS5000225: This tenant has been blocked due to inactivity`, your MSA's auto-created "Default Directory" tenant has been deactivated and the sign-in dropdown offers only that dead tenant — so you can't register an app there. You need a **live** tenant instead. Two options:
>
> 1. **Microsoft 365 Developer Program** — https://developer.microsoft.com/en-us/microsoft-365/dev-program provisions a fresh sandbox tenant. (Eligibility was tightened in 2024, so instant free signup isn't guaranteed for every account.)
> 2. **A fresh Microsoft account** — a brand-new `@outlook.com` gets its own unblocked tenant. Register the app there as "any org directory **and** personal Microsoft accounts," then sign in at `whoami` time with whichever account owns the mailbox you actually want to read (the multi-tenant registration accepts either).

## Setup (from source, for development)

If you want to hack on the server rather than just install it, clone and build locally:

```bash
git clone https://github.com/jaingxyz/personal-outlook-mcp.git
cd personal-outlook-mcp

cp .env.example .env
# edit .env and paste your AZURE_CLIENT_ID

npm ci            # exact versions from package-lock.json
npm run build
npm test          # unit + integration tests, no network
```

### Available scripts

| Script                  | What it does                                            |
| ----------------------- | ------------------------------------------------------- |
| `npm run build`         | Compile TypeScript to `dist/`.                          |
| `npm run dev`           | `tsc --watch` for iterative development.                |
| `npm run start`         | Launch the compiled MCP server (only useful via stdio). |
| `npm run whoami`        | Auth smoke test; prints `/me` from Graph.               |
| `npm test`              | Unit tests (Graph mocked, no network).                  |
| `npm run test:watch`    | Vitest in watch mode.                                   |
| `npm run test:coverage` | Coverage report under `coverage/`.                      |
| `npm run lint`          | ESLint over `src/` and `test/`.                         |
| `npm run lint:fix`      | Lint + auto-fix.                                        |
| `npm run format`        | Prettier write across the repo.                         |
| `npm run format:check`  | Prettier check (CI uses this).                          |
| `npm run clean`         | Remove `dist/` and `coverage/`.                         |

## First run (device-code auth)

```bash
npm run whoami
```

The script will print to **stderr**:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Open the URL, enter the code, sign in with your personal Microsoft account, grant the requested permissions. The script then prints your account info as JSON on stdout. Tokens are persisted to the OS keyring (via [`@napi-rs/keyring`](https://github.com/Brooooooklyn/keyring-node)) under service `personal-outlook-mcp`, account `msal-cache` — subsequent runs refresh silently.

To sign out and forget the cached token:

```bash
node -e "import('./dist/auth.js').then(m => m.signOut())"
```

## MCPB bundle (Claude Desktop one-click install)

A pre-packaged [MCPB bundle](https://github.com/anthropics/mcpb) is built from the `npm run bundle` script. **Currently macOS arm64 only** — the bundle includes a platform-specific native binary for the OS keyring backend, and we ship one platform at a time. Other OSes should install via `npx -y @jaingxyz/personal-outlook-mcp` (which fetches the right native binary at install time).

To build and test locally:

```bash
npm run bundle
# produces personal-outlook-mcp.mcpb (~5 MB)
# double-click in Finder to install into Claude Desktop
```

The bundle's manifest declares three user-config fields, prompted at install: Azure client ID, tenant (default `consumers`), display timezone (default `America/Los_Angeles`).

## Using from a local source clone

If you cloned the repo (instead of installing via npx), point Claude Desktop at the built script:

```json
{
  "mcpServers": {
    "personal-outlook": {
      "command": "node",
      "args": ["/absolute/path/to/personal-outlook-mcp/dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "YOUR-CLIENT-ID-HERE",
        "AZURE_TENANT": "consumers"
      }
    }
  }
}
```

> First-run auth from inside Claude Desktop is awkward because the device-code prompt is written to stderr, which Claude Desktop doesn't surface. Run `npm run whoami` from a terminal **once** before launching Claude Desktop to seed the keychain — after that, the MCP server picks up the cached token silently.

## Tools

| Tool                                  | Purpose                                                                                      |
| ------------------------------------- | -------------------------------------------------------------------------------------------- |
| `personal_email_list_folders`         | List mail folders with id, display name, unread/total counts.                                |
| `personal_email_list_recent`          | Newest-first messages in a folder. Supports `unreadOnly`.                                    |
| `personal_email_search`               | Free-text or KQL search across the mailbox. Ranked by relevance.                             |
| `personal_email_read`                 | Fetch one message with full body (text or html).                                             |
| `personal_email_mark_read`            | Mark a message read or unread.                                                               |
| `personal_email_move`                 | Move a message to another folder (id or well-known name).                                    |
| `personal_email_delete`               | Soft-delete (move to Deleted Items) by default; `hardDelete=true` is unrecoverable.          |
| `personal_email_send`                 | Send a new email immediately. Saves a copy to Sent Items.                                    |
| `personal_email_reply`                | Reply to a message by id. `replyAll=true` to reply to all recipients.                        |
| `personal_email_create_draft`         | Create a draft in Drafts without sending. Returns a draftId.                                 |
| `personal_email_send_draft`           | Send a previously created draft by id.                                                       |
| `personal_email_list_attachments`     | List attachments on a message (id, name, contentType, size).                                 |
| `personal_email_download_attachment`  | Save a file attachment to disk. Defaults to `~/Downloads/personal-outlook-mcp/`.             |
| `personal_calendar_list_calendars`    | List calendars (primary, holidays, custom group calendars).                                  |
| `personal_calendar_list_events`       | Events in a date range (calendarView; recurring series expanded into occurrences).           |
| `personal_calendar_read_event`        | Full event details: attendees, body, recurrence pattern.                                     |
| `personal_calendar_create_event`      | Create an event. If attendees are listed, Graph sends invites.                               |
| `personal_calendar_update_event`      | Update subject/time/location/body/attendees. Refuses single occurrences of recurring series. |
| `personal_calendar_cancel_event`      | Cancel (sends notice) or hard-delete (no notice).                                            |
| `personal_calendar_respond_to_invite` | accept / tentativelyAccept / decline an invite.                                              |

Folder names accept Outlook well-known names: `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`. Custom folders need their id (get it from `personal_email_list_folders`).

Calendar event times use Graph's native shape: `{ "dateTime": "2026-05-20T15:00:00", "timeZone": "America/Los_Angeles" }`. The `dateTime` field is a local datetime with no offset; `timeZone` is an IANA name. Output times default to `America/Los_Angeles`; override with the `PERSONAL_OUTLOOK_TZ` env var.

## Troubleshooting

- **`Missing required env var AZURE_CLIENT_ID`** — `.env` not present or not loaded. Check `cat .env`.
- **`AuthError: invalid_grant` thrown immediately, with no device code printed** — the app registration is misconfigured, not your sign-in. The `consumers` authority only accepts personal Microsoft accounts via a public client. Two causes: (1) **Supported account types** doesn't include personal Microsoft accounts — set it to "any org directory **and** personal Microsoft accounts" (manifest `signInAudience: AzureADandPersonalMicrosoftAccount`); (2) **Allow public client flows** is not **Yes**. Both live on the app's **Authentication** blade — see "Azure setup" steps 2 and 5. Fix and re-run `npm run whoami`; you should now get a real device code.
- **`AADSTS5000225`** during sign-in (or while trying to register the app) — the tenant your app registration lives in has been deactivated for inactivity, and the portal's tenant dropdown offers only that dead tenant. You need a live tenant: either the [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) sandbox (subject to eligibility) or a fresh `@outlook.com` account (whose new tenant isn't blocked). Register the app as multi-tenant + personal accounts there, then sign in with whichever account owns the mailbox you want. See "Azure setup" for the full note.
- **MCP client says "server crashed" with no useful output** — common cause is something writing to stdout, which corrupts the JSON-RPC stream. All non-protocol output must go to stderr.
- **`Re-authentication required: ...`** at runtime — the cached token can't be refreshed (scopes changed, password changed, refresh expired). Run `npm run whoami` from a terminal to do device-code flow once; the MCP server picks up the new token automatically on the next tool call.

## Architecture

See [CLAUDE.md](./CLAUDE.md).

## License

AGPL-3.0-or-later. See [LICENSE](./LICENSE) for the full text.

This is a copyleft license that closes the SaaS loophole: anyone running a modified version as a network service must release their modifications under the same terms. Suitable for personal projects that want to remain free as in freedom while preventing closed-source forks from being hosted as commercial products.

## Acknowledgements

Built with assistance from [Claude](https://www.anthropic.com/claude) (Anthropic). The architecture, design decisions, and final code review remain the responsibility of the human author; Claude served as a pair-programming and drafting collaborator.
