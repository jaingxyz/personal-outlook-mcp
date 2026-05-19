# personal-outlook-mcp

A [Model Context Protocol](https://modelcontextprotocol.io/) server that exposes a personal Outlook (consumer Microsoft account) inbox to MCP clients like Claude Desktop. Talks to Microsoft Graph over HTTPS, uses MSAL device-code flow for OAuth, and stores tokens in the macOS Keychain.

## Prerequisites

- Node.js 20+
- A client ID for an Azure AD app — see "Auth setup" below
- macOS (token cache uses the system keychain via `keytar`)

## Auth setup

You need an Azure AD app registration's **Application (client) ID**. The recommended path is to register your own (you control the consent screen, rate limits, and scope set), but for getting started today you can borrow Microsoft's Graph Explorer client ID.

### Option A — Register your own app (recommended)

1. Go to https://portal.azure.com → "App registrations" → "New registration".
2. **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts".
3. Leave the redirect URI blank. Click Register.
4. In the new app: **Authentication** → "Add a platform" → "Mobile and desktop applications" → add the redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`.
5. Same page, scroll to **Advanced settings** → set "Allow public client flows" = **Yes**. Save.
6. Copy the **Application (client) ID** from the Overview page.

The required Graph delegated permissions (`Mail.ReadWrite`, `Mail.Send`, `offline_access`, `User.Read`) are requested at sign-in. For personal accounts the user consents at the device-code prompt — no admin consent required.

> If `portal.azure.com` returns `AADSTS5000225: This tenant has been blocked due to inactivity`, your MSA's auto-created "Default Directory" tenant has been deactivated and you can't reach the portal. Either join the [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) for a fresh sandbox tenant, or use Option B below.

### Option B — Borrow the Graph Explorer client ID (dev-only)

Skip the Azure portal entirely and use Microsoft's public Graph Explorer client ID:

```
AZURE_CLIENT_ID=REDACTED-PUBLIC-CLIENT-ID
AZURE_TENANT=consumers
```

Caveats: the consent screen will say "Microsoft Graph Explorer" instead of your app's name; rate limits are shared with anyone else doing the same thing; Microsoft can revoke or restrict it. Fine for personal use, not appropriate for anything you'd ship.

## Setup

```bash
cp .env.example .env
# edit .env and paste your AZURE_CLIENT_ID

npm install
npm run build
```

## First run (device-code auth)

```bash
npm run whoami
```

The script will print to **stderr**:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Open the URL, enter the code, sign in with your personal Microsoft account, grant the requested permissions. The script then prints your account info as JSON on stdout. Tokens are persisted to the macOS Keychain under service `personal-outlook-mcp`, account `msal-cache` — subsequent runs refresh silently.

To sign out and forget the cached token:

```bash
node -e "import('./dist/auth.js').then(m => m.signOut())"
```

## Using with Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` and add:

```json
{
  "mcpServers": {
    "personal-outlook": {
      "command": "node",
      "args": ["/Users/YOU/code/personal_outlook_mcp/dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "YOUR-CLIENT-ID-HERE",
        "AZURE_TENANT": "consumers"
      }
    }
  }
}
```

Substitute the absolute path to `dist/index.js` and your client ID. Restart Claude Desktop. The tools below should appear in the tool picker prefixed with `personal_email_`.

> First-run auth from inside Claude Desktop is awkward because the device-code prompt is written to stderr, which Claude Desktop doesn't surface. Run `npm run whoami` from a terminal **once** before launching Claude Desktop to seed the keychain — after that, the MCP server picks up the cached token silently.

## Tools

| Tool | Purpose |
|---|---|
| `personal_email_list_folders` | List mail folders with id, display name, unread/total counts. |
| `personal_email_list_recent` | Newest-first messages in a folder. Supports `unreadOnly`. |
| `personal_email_search` | Free-text or KQL search across the mailbox. Ranked by relevance. |
| `personal_email_read` | Fetch one message with full body (text or html). |
| `personal_email_mark_read` | Mark a message read or unread. |
| `personal_email_move` | Move a message to another folder (id or well-known name). |
| `personal_email_delete` | Soft-delete (move to Deleted Items) by default; `hardDelete=true` is unrecoverable. |

Folder names accept Outlook well-known names: `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`. Custom folders need their id (get it from `personal_email_list_folders`).

## Troubleshooting

- **`Missing required env var AZURE_CLIENT_ID`** — `.env` not present or not loaded. Check `cat .env`.
- **`AADSTS5000225`** during sign-in — the tenant the client ID lives in is deactivated. Re-register the app in a different tenant, or switch to Option B.
- **MCP client says "server crashed" with no useful output** — common cause is something writing to stdout, which corrupts the JSON-RPC stream. All non-protocol output must go to stderr.
- **`InteractionRequiredAuthError`** at runtime after working before — the cached refresh token expired or your password changed. Run `npm run whoami` to re-auth.

## Architecture

See [CLAUDE.md](./CLAUDE.md).

## License

MIT
