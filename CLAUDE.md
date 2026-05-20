# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An MCP (Model Context Protocol) server, written in TypeScript, that exposes a single user's personal Outlook account (consumer Microsoft account, not work/school) to MCP clients via Microsoft Graph. Transport is **stdio only** — there is no HTTP server.

## Commands

- `npm install` — install deps
- `npm run build` — compile `src/` to `dist/` via `tsc`
- `npm run dev` — `tsc --watch`
- `npm run start` — run the compiled MCP server (`node dist/index.js`); only useful when launched by an MCP client over stdio
- `npm run whoami` — auth smoke test: runs device-code flow if needed and prints `/me` from Graph. Use this first when debugging auth.
- `npm run clean` — remove `dist/`

There is no test runner wired up yet.

## Architecture

```
src/
  index.ts          entry — boots stdio MCP server
  server.ts         registers MCP tools (Zod schemas -> Graph wrapper)
  auth.ts           MSAL PublicClientApplication + keychain-backed token cache
  graph.ts          thin Microsoft Graph wrapper (mail ops only)
  tools/            one module per tool group: read, mutate, send
  scripts/whoami.ts auth smoke test
```

### Critical invariants

- **stdout is reserved for the MCP JSON-RPC stream.** Any logging — including the device-code prompt MSAL emits on first run — MUST go to stderr (`console.error`). A stray `console.log` will corrupt the protocol stream and the client will disconnect with cryptic errors.
- **Token cache lives in the macOS Keychain** via `keytar`, service `personal-outlook-mcp`, account `msal-cache`. We implement MSAL's `ICachePlugin` to read/write the serialized cache. Do not write tokens to disk.
- **Auth flow is device code** (`acquireTokenByDeviceCode`), but **gated**. `getAccessToken({ interactive: true })` is the only path that triggers device-code; everything else (the MCP server, every tool call) passes `interactive: false` (the default) and gets `ReauthRequiredError` if silent refresh fails. This prevents the MCP server from hanging on a device-code prompt that Claude Desktop can't surface. Only `src/scripts/whoami.ts` runs in interactive mode; it's the user's escape hatch when scopes change or the cache is stale.
- **Tenant is `consumers`** (personal accounts only). Switching to `common` admits work/school accounts and changes which scopes require admin consent.
- **`/me.mail` is `null` for personal MSAs.** Use `userPrincipalName` if you need the email-shaped identifier of the signed-in user. Don't add code that depends on `/me.mail` being populated.

### Graph specifics worth remembering

- Required delegated scopes: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `offline_access`, `User.Read`. `offline_access` is what gives us a refresh token — without it, the user re-auths every hour. Adding a new scope invalidates the cached refresh token; users have to redo device-code consent (run `npm run whoami` to seed the new scope).
- Search uses Graph's `$search` query parameter (KQL-style). When `$search` is present, Graph forbids `$orderby` — code paths that combine them must branch.
- Folder operations accept either a folder id or a well-known name (`inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`).
- Delete-by-default semantics: our `delete` tool moves to `deleteditems` rather than hard-deleting, to match Outlook's UI behavior. Hard delete is a separate flag.
- `POST /me/messages/{id}/move` returns the message with its **new id** in the destination folder. The pre-move id no longer resolves. Tools that move a message return both for caller convenience.
- Calendar event times use `{ dateTime, timeZone }` where `dateTime` is **local form, no offset** and `timeZone` is an IANA name. Don't try to combine them into an ISO offset string before sending — Graph rejects offset strings in this shape.
- `calendarView` (start/end query params) expands recurring series into individual occurrences. `/me/events` returns the series master + standalone events; use `calendarView` for "what's on my calendar this week" questions.
- Updating a single occurrence of a recurring series via `PATCH /me/events/{id}` is a known Graph footgun. We refuse it explicitly: `update_event` checks `type === "occurrence" | "exception"` and errors. Series-master edits are fine.
- `Prefer: outlook.timezone="..."` header makes Graph render output times in that zone. Apply it to GET endpoints that read events. Set zone via `PERSONAL_OUTLOOK_TZ` env var; defaults to `America/Los_Angeles`.

### Tool naming

All tools are prefixed `personal_email_*` so a future multi-account setup can coexist (e.g., `work_email_*`). Keep the prefix when adding tools.

## Claude Desktop integration

The MCP server is launched by Claude Desktop via `claude_desktop_config.json` with stdio. Important quirks:

- Claude Desktop swallows stderr — first-run device-code prompts won't be visible to the user inside the app. Workflow is: run `npm run whoami` from a terminal once to seed the keychain, then launch Claude Desktop. Future runs use silent token refresh.
- Tools that mutate state (`mark_read`, `move`, `delete`) carry MCP `annotations.destructiveHint`, which Claude Desktop uses to show an extra confirmation dialog. Don't strip these annotations.

## Configuration

`.env` (gitignored) supplies:
- `AZURE_CLIENT_ID` — Application (client) ID from the Azure AD app registration
- `AZURE_TENANT` — defaults to `consumers`; change only with intent

`.env.example` is the template. Never commit `.env`.

## References

- MCP spec: https://modelcontextprotocol.io/docs
- MCP TypeScript SDK: https://github.com/modelcontextprotocol/typescript-sdk
- Graph mail API: https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
- Device code flow: https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code
