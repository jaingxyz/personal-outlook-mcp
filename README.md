# personal-outlook-mcp

A [Model Context Protocol](https://modelcontextprotocol.io/) server that exposes a personal Outlook (consumer Microsoft account) inbox to MCP clients like Claude Desktop. Talks to Microsoft Graph over HTTPS, uses MSAL device-code flow for OAuth, and stores tokens in the macOS Keychain.

Status: in progress (Phase 1: scaffolding).

## Prerequisites

- Node.js 20+
- An Azure AD app registration (free) — see "Azure setup" below
- macOS (token cache uses the system keychain via `keytar`)

## Azure setup

1. Go to https://portal.azure.com -> "App registrations" -> "New registration".
2. **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts".
3. Leave the redirect URI blank for now. Click Register.
4. In the new app: **Authentication** -> "Add a platform" -> "Mobile and desktop applications" -> add the redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`.
5. Same page, scroll to **Advanced settings** -> set "Allow public client flows" = **Yes**. Save.
6. Copy the **Application (client) ID** from the Overview page.

The required Graph delegated permissions (`Mail.ReadWrite`, `Mail.Send`, `offline_access`, `User.Read`) are requested at sign-in; for personal accounts the user consents at the device-code prompt — no admin consent required.

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

The script will print something like:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Open the URL, enter the code, sign in with your personal Microsoft account, grant the requested permissions. The script then prints your account info. Tokens are persisted to the macOS Keychain under service `personal-outlook-mcp`.

## Using with Claude Desktop

(Filled in at Phase 5.)

## Tools

(Filled in as each phase lands.)

## License

MIT
