# Security Policy

## Reporting a vulnerability

Please do **not** open a public issue for security vulnerabilities.

Use GitHub's [private vulnerability reporting](https://github.com/jaingxyz/personal-outlook-mcp/security/advisories/new) on this repository instead. That channel notifies the maintainer privately and creates a draft advisory.

You can expect an initial response within ~7 days. Fix timelines depend on severity and reachability — this is a personal project, not a service.

## Scope

In scope:

- Code in `src/` that handles authentication, token storage, message contents, file writes, or input validation.
- Dependency vulnerabilities flagged by Dependabot or `npm audit`.

Out of scope:

- Vulnerabilities in Microsoft Graph itself (report to Microsoft).
- Vulnerabilities in Claude Desktop or other MCP clients (report to those vendors).
- Misconfigurations of the user's Azure AD app registration.

## Supported versions

Only the latest commit on `main` is supported. There are no maintained release branches.
