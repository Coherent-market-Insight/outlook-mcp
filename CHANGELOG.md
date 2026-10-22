# Changelog

All notable changes are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com), and the project uses [Semantic Versioning](https://semver.org).

## [1.0.0] - 2026-10-22

First release.

### Added
- MCP server for the Outlook desktop client on Windows, written in C# on .NET 10, calling Outlook's COM object model in-process and shipped as one self-contained executable.
- 31 tools: mail (list and search, read, thread, send, drafts, reply, forward, flag, move, delete, attachments, folders), calendar (events with recurrence, create, update, delete, respond, free/busy, calendars), contacts and address-list lookup, tasks, and `status`.
- Multi-account support: send as any account, list one mailbox or all, per-account calendars, contacts and tasks.
- Token-lean output: short ids, one line per item, cleaned bodies with continuation offsets, optional previews, batch operations.
- Safety controls: `--read-only`, `--disable-send`, `--toolsets`, `--disable-tools`, `--allowed-domains`, draft-first replies and forwards, fail-closed sends.
- `--debug` and `--log-file` with argument redaction; a watchdog that turns a blocked Outlook into a clear error and drops abandoned work.
- `doctor`, `install --client <name>` for seven MCP clients, `install-skill`, and `tools --markdown`.
- Bundled agent skill and a live integration suite.
