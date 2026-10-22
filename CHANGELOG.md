# Changelog

All notable changes are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com), and the project uses [Semantic Versioning](https://semver.org). The section for a version is used as that release's notes.

## [1.0.0] - 2026-10-22

First release.

### Platform
- MCP server for the **Outlook desktop client on Windows**, written in C# on .NET 10. It calls Outlook's COM object model directly, in-process, and ships as **one self-contained 39 MB executable**: no Node, Python, PowerShell or .NET install is needed.
- No cloud registration, OAuth or Graph API. Anything you can open in Outlook works, including **POP and IMAP accounts**.
- All COM work runs on one dedicated STA thread; the connection is created lazily and recovers once from Outlook being closed or restarted.
- Fast: connects in about 0.2 to 0.4 s and reads a 1,300-message inbox in about 30 ms. About 77 MB idle, about 130 MB after heavy use.

### Mail (14 tools)
- `mail_list`: list **and** search in one tool, newest first, with server-side filters for text, sender, recipient, subject, unread, attachments, flagged, category, date range; paging with an exact `offset`; optional previews; any folder by name, path or id; `account="all"` to merge the same folder across every mailbox.
- `mail_get`: headers, attachment list and a cleaned body, capped with an exact continuation offset; optional full headers and addresses; never marks a message read.
- `mail_thread`: a whole conversation across Inbox, Sent and Drafts with trimmed bodies.
- `mail_send`, `mail_draft_save`, `mail_draft_send`: plain or HTML, cc and bcc, attachments, importance, deferred delivery, send as a chosen account. Saved drafts can be edited in place.
- `mail_reply` and `mail_forward`: reply, reply-all and forward with your text above the quoted original and the original formatting preserved. They **save a draft unless `send=true`**.
- `mail_update`: read state, follow-up flag, categories (set, add, remove) and importance, on many messages in one call.
- `mail_move`, `mail_delete` (to Deleted Items, or permanent), `mail_attachments` (list, or save to disk and return paths), `mail_folders` (tree with unread and total counts), `mail_folder_create`.

### Calendar (8 tools)
- `calendar_events`: time-ordered events in a range with **recurring meetings expanded**, across any calendar.
- `calendar_event_get`: full details including organiser, attendees with their responses, recurrence and reminder.
- `calendar_event_create`: all-day events, reminders, show-as status, categories, attendees (required and optional), and simple recurrence (daily, weekly with days, monthly, yearly; until a date or a count). Meetings are **saved unsent unless `send_invites=true`**.
- `calendar_event_update` and `calendar_event_delete`: change or remove a whole event or **a single occurrence** of a series; a start change keeps the duration; optional attendee notification or cancellation.
- `calendar_respond`: accept, tentatively accept or decline a meeting request, with an optional note.
- `calendar_free_busy`: free windows in working hours from your calendar, plus other people's availability on Exchange and Microsoft 365.
- `calendar_list`: every calendar in the profile.

### Contacts (5 tools) and tasks (3 tools)
- `contacts_search`, `contact_get`, `contact_save` (create or update), `contact_delete`, and `recipient_resolve`, which also searches the company address list on Exchange.
- `tasks_list` (open, completed or all, with due-date and text filters), `task_save` (create, edit, set status, percent, reminder, complete), `task_delete`.

### Multiple accounts
- Every account is addressable by SMTP address or display name through `account`, and each has its own calendar, contacts and tasks.
- **Sending as an account works for POP and IMAP**: Outlook ignores or rejects `SendUsingAccount` there, so the message is created inside the target account's own store, and checked to be in the right store before it is sent. Verified end to end.
- Replies go out from the mailbox the original message arrived in.

### Token efficiency
- Tool definitions cost about 6.6K tokens for all 31 tools, about 2.4K in `--read-only` mode and about 1.1K for mail read-only.
- Short ids (`7ecb73:E4362200`, 15 characters instead of 55), one line per item, no JSON envelope.
- Bodies are cleaned: link and `mailto:` noise, gateway banners, the quoted reply chain, trailing legal footers and double spacing are removed.
- Batch operations, thread view, attachments returned as paths, and previews that are opt-in.
- A unit test fails the build if the tool surface grows past its budget.

### Safety
- `--read-only`, `--disable-send`, `--toolsets`, `--disable-tools` and `--allowed-domains`, enforced in the server rather than in a prompt.
- MCP annotations (`readOnlyHint`, `destructiveHint`, `openWorldHint`) on every tool.
- Fail-closed sending: every recipient must resolve, and pass the allow-list when one is set.
- When Outlook is blocked by a dialog, the call returns a clear error after a timeout, queued work is dropped so a stale send can never fire later, and new calls fail fast.
- Errors are returned as normal MCP error results with a message written for the model; no stack traces leak.

### Diagnostics and debugging
- `outlook-mcp doctor`: Windows version, COM registration, connection time, accounts with their types, and inbox read speed.
- `--debug` and `--log-file`: one line per call with tool, timing and size. Free-text arguments are logged only by length, so logs are safe to share.
- `status` tool: Outlook version, profile, accounts and the active policy.

### Command line and installers
- `install --client <name>` for **Claude Code, Claude Desktop, Cursor, VS Code, Gemini CLI, OpenCode and Codex CLI**. It preserves everything else in the client's config, tolerates comments and trailing commas, writes a `.bak-outlook-mcp` backup, updates in place when re-run, and supports a second Claude Code profile with `--config-dir`. Server flags can be baked in after `--`.
- `install-skill`: installs the bundled agent skill, which teaches a token-lean, draft-first workflow.
- `tools --markdown`: lists the tools under the current flags and generates the tool reference.

### Quality
- 57 unit and installer tests, run on every push on Windows.
- A live integration suite (75+ checks) that runs every tool against a real Outlook profile, creates only tagged test items, and removes only what it created, by id.
- Continuous integration builds, tests, publishes the single-file executable and smoke-tests it; the release workflow verifies that the built file reports the tagged version and publishes checksums.
- Documentation: README, architecture, installation, token efficiency, troubleshooting, tool reference, security policy, contributing guide and code of conduct.

### Known limitations
- Classic Outlook only; the new Outlook for Windows has no COM interface.
- The executable is not code-signed, so Windows SmartScreen may warn on first run. Verify downloads with the published `.sha256` files.
- No automatic signature, and no mail rules or out-of-office, because Outlook's object model does not expose them.
- Other people's free/busy needs Exchange or Microsoft 365.
- `calendar_respond` and sending meeting invitations are implemented but have not been exercised against live invitations.
