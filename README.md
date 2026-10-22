# Outlook MCP

[![CI](https://github.com/Coherent-market-Insight/outlook-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/Coherent-market-Insight/outlook-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078d4)
![.NET 10](https://img.shields.io/badge/.NET-10-512bd4)

A [Model Context Protocol](https://modelcontextprotocol.io) server for the **Outlook desktop client on Windows**. It gives an MCP client mail, calendar, contacts and tasks across every account in your Outlook profile, and it is built to spend as few tokens as possible.

It talks to Outlook's COM object model directly from one self-contained `.exe`. There is no cloud app registration, no OAuth, no Graph API and no password handling: if you can open the account in Outlook, the server can use it. That includes POP and IMAP accounts that API-based connectors cannot reach.

## Why this one

| | |
|---|---|
| **One file, in-process** | A single 39 MB executable. No Node, no Python, no PowerShell. Connects to Outlook in about 0.3 s and reads a 1,300-message inbox in about 30 ms. |
| **Token-lean by design** | Short ids (`7ecb73:E4362200`), one line per message, cleaned bodies (link noise, quoted reply chains, legal footers and gateway banners removed), batch operations, optional previews. See [docs/TOKEN-EFFICIENCY.md](docs/TOKEN-EFFICIENCY.md). |
| **Safe by default** | Replies and forwards save drafts unless you ask to send. `--read-only`, `--disable-send` and `--allowed-domains` are enforced in the server, not in a prompt. Unresolved recipients fail closed. |
| **Real multi-account** | Send as any account, list one mailbox or all of them at once, and address each account's calendar, contacts and tasks. |
| **Full coverage** | 31 tools: mail, folders, attachments, drafts, threads, calendar with recurrence and free/busy, meeting responses, contacts, the address list, tasks. |
| **Debuggable** | `outlook-mcp doctor`, a per-call debug log that never records message text, and a watchdog that turns a blocked Outlook into a clear error instead of a hang. |

## Requirements

- Windows 10 or 11
- **Classic Outlook** (Microsoft 365 Apps, Outlook 2016/2019/2021). The *new Outlook for Windows* has no COM interface and cannot be used.
- Nothing else to run the released executable. Building from source needs the .NET 10 SDK.

## Install

1. Download `outlook-mcp.exe` from the [latest release](https://github.com/Coherent-market-Insight/outlook-mcp/releases/latest) and put it somewhere permanent, for example `%LOCALAPPDATA%\outlook-mcp\`.
2. Check that it can see your profile:

   ```powershell
   .\outlook-mcp.exe doctor
   ```

3. Register it with your MCP client. The installer edits the client's config for you, keeps a `.bak-outlook-mcp` backup, and is safe to run again:

   ```powershell
   .\outlook-mcp.exe install --client claude-code      # claude-desktop | cursor | vscode | gemini | opencode | codex
   ```

4. Optional: install the bundled skill, which teaches an agent the token-lean workflow:

   ```powershell
   .\outlook-mcp.exe install-skill
   ```

Per-client details, manual JSON/TOML snippets and read-only presets are in [docs/INSTALL.md](docs/INSTALL.md).

## Configuration

Flags and environment variables are equivalent. Pass flags after `--` when installing, for example `install --client cursor -- --read-only`.

| Flag | Environment variable | Effect |
|---|---|---|
| `--read-only` | `OUTLOOK_MCP_READ_ONLY=1` | Expose only read tools (13 of 31). |
| `--disable-send` | `OUTLOOK_MCP_DISABLE_SEND=1` | Hide every tool that can send mail or meeting responses. |
| `--toolsets mail,calendar` | `OUTLOOK_MCP_TOOLSETS` | Limit to `mail`, `calendar`, `contacts`, `tasks`. `system` is always on. |
| `--disable-tools a,b` | `OUTLOOK_MCP_DISABLED_TOOLS` | Hide individual tools by name. |
| `--allowed-domains x.com,y.com` | `OUTLOOK_MCP_ALLOWED_DOMAINS` | Refuse to send to any other domain. |
| `--debug`, `--log-file <path>` | `OUTLOOK_MCP_DEBUG=1`, `OUTLOOK_MCP_LOG_FILE` | Log one line per call: tool, timing, size. Never message text. |
| | `OUTLOOK_MCP_TIMEOUT_SECONDS` | How long to wait for a blocked Outlook (default 60). |
| | `OUTLOOK_MCP_BODY_MAX_CHARS` | Default body length returned by `mail_get` (default 3000). |

## What it can do

| Area | Tools |
|---|---|
| **Mail** (14) | `mail_list` (list and search with filters, paging, `account="all"`), `mail_get`, `mail_thread`, `mail_send`, `mail_draft_save`, `mail_draft_send`, `mail_reply`, `mail_forward`, `mail_update` (read state, follow-up flag, categories, importance, in batches), `mail_move`, `mail_delete`, `mail_attachments`, `mail_folders`, `mail_folder_create` |
| **Calendar** (8) | `calendar_list`, `calendar_events` (recurrences expanded), `calendar_event_get`, `calendar_event_create` (all-day, reminders, show-as, attendees, recurrence), `calendar_event_update`, `calendar_event_delete` (single occurrence or whole series, optional cancellations), `calendar_respond`, `calendar_free_busy` |
| **Contacts** (5) | `contacts_search`, `contact_get`, `contact_save`, `contact_delete`, `recipient_resolve` (saved contacts and the company address list) |
| **Tasks** (3) | `tasks_list`, `task_save` (create, edit, complete), `task_delete` |
| **System** (1) | `status` |

The full table with access levels is generated in [docs/TOOLS.md](docs/TOOLS.md).

Example prompts that work well:

- "Which unread mail from the last two days needs a reply? Draft replies, don't send."
- "Find a 45-minute slot next week when both of us are free and create the meeting without sending invites."
- "Move everything from this sender into Inbox/Clients, then flag the newest one for follow-up."
- "Send this as the sales mailbox, to the test address only."

## Multiple accounts

Every account in the profile is addressable by SMTP address or display name through the `account` argument.

- **Sending as an account.** Outlook ignores or rejects `SendUsingAccount` for POP and IMAP accounts. The server therefore creates the message inside the target account's own store, which is the mechanism Outlook honours. The message is checked to be in the right store before it is sent; otherwise nothing is sent.
- **Listing.** `mail_list` with `account="all"` merges the same folder across every mailbox, newest first.
- **Calendars, contacts, tasks.** Each takes an account or calendar name; `calendar_list` shows what exists.
- **Replies** go out from the mailbox the original message arrived in.

## Safety model

The protections live in the server so that a mistaken or manipulated prompt cannot bypass them.

- Tools carry MCP annotations (`readOnlyHint`, `destructiveHint`), so clients can gate them.
- `mail_reply`, `mail_forward` and `calendar_event_create` with attendees only **save** unless you pass `send=true` / `send_invites=true`.
- Before any send, all recipients must resolve, and, when configured, every domain must be on the allow-list. A failure sends nothing.
- A call that cannot complete because Outlook is blocked is abandoned, and queued work behind it is dropped so a stale send can never fire later.
- Message bodies, subjects and addresses are never written to the debug log.

See [SECURITY.md](SECURITY.md) for the threat model and how to report a vulnerability.

## Token efficiency

Measured on this repository's build (about 4 characters per token):

| Configuration | Tools | Definition size |
|---|---|---|
| Everything | 31 | ~6.6K tokens |
| `--read-only` | 13 | ~2.4K tokens |
| `--toolsets mail` | 15 | ~3.1K tokens |
| `--toolsets mail --read-only` | 5 | ~1.1K tokens |

Per call: an inbox line is about 25 tokens; ids are about 15 characters instead of 55; `mail_get` returns a cleaned, capped body with an exact continuation offset. Details and tuning advice are in [docs/TOKEN-EFFICIENCY.md](docs/TOKEN-EFFICIENCY.md).

## Limitations

- Classic Outlook on Windows only.
- No automatic signature (Outlook's object model does not apply one to programmatic mail); include it in the body.
- Mail rules, out-of-office and delegate access are not exposed by Outlook's object model.
- Other people's free/busy needs Exchange or Microsoft 365; other account types report no data.
- Outlook may show a security prompt the first time another program drives it, if antivirus is not detected. Click through once, or see [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

## Development

```powershell
dotnet build
dotnet test tests/OutlookMcp.Tests              # unit and installer tests, no Outlook needed
dotnet publish src/OutlookMcp -c Release -r win-x64 -p:PublishSingleFile=true -o dist

# live suite against your own Outlook profile (Node 18+); read the safety notes in the file first
node scripts/live-test.mjs
```

The live suite creates tagged test items and removes them again. It only sends mail when `OUTLOOK_MCP_TEST_SEND_TO` is set, and only to that address. Architecture notes are in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md); contribution guidelines in [CONTRIBUTING.md](CONTRIBUTING.md).

## License

[MIT](LICENSE) © Coherent Market Insights Pvt. Ltd.
