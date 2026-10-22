# Security

## Reporting a vulnerability

Please do not open a public issue for a security problem. Use GitHub's private reporting instead: **Security → Report a vulnerability** on this repository. Include the version (`outlook-mcp version`), what you did, and what you expected. We aim to acknowledge reports within three working days.

## What this server can do

Outlook MCP runs on your machine, as your Windows user, and drives your own Outlook profile through COM. It can read and change everything that profile can: mail, calendars, contacts, tasks, and it can send mail as any account in the profile. It has no network listener (stdio only), stores no credentials, and calls no external service.

That makes the practical risk **prompt injection**: text inside a message, invitation or contact that tries to make a connected agent read, forward or send something it should not. The server cannot tell instructions from content, so it limits what is possible instead.

## Controls built into the server

| Control | Effect |
|---|---|
| `--read-only` | Removes every tool that changes anything. Use it for triage and reporting agents. |
| `--disable-send` | Removes every tool that can send mail or meeting responses. |
| `--toolsets` / `--disable-tools` | Expose only what a job needs. |
| `--allowed-domains` | A send to any other domain is refused before it happens. |
| Draft-first | `mail_reply` and `mail_forward` save a draft unless `send=true`. Meetings with attendees are saved unsent unless `send_invites=true`. |
| Fail-closed sends | Every recipient must resolve to an address, and, when configured, pass the allow-list. Otherwise nothing is sent. |
| Account check | A send as a specific account only proceeds if the message was created in that account's store. |
| Abandoned-call drop | If Outlook is blocked and a call times out, queued work is dropped, so an old send cannot fire later. |
| MCP annotations | Tools declare `readOnlyHint` and `destructiveHint`, so clients can prompt before risky calls. |
| Log hygiene | Debug logs record tool name, timing and sizes, and reduce argument text to its length. Subjects, bodies and addresses are never logged. |

## Recommendations

- Give agents that read untrusted mail `--read-only`, or at least `--disable-send`.
- If an agent must send, set `--allowed-domains` to the domains it legitimately writes to.
- Keep your client's per-tool approval on for tools marked destructive.
- Saved attachments land in `%TEMP%\outlook-mcp`; treat them as untrusted files.

## Supported versions

Security fixes go into the latest release.
