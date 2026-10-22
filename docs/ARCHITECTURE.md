# Architecture

```
MCP client ──stdio──▶ Host (Microsoft.Extensions.Hosting + MCP SDK)
                         │  call-tool filter: debug log, error → isError result
                         ▼
                      Registry ── [ToolMeta] methods ── policy (toolsets, read-only, send)
                         │
                         ▼
                      Tool methods (static, async)
                         │  Sta.Run(...)
                         ▼
                 dedicated STA thread ──▶ Ol (late-bound COM) ──▶ Outlook.Application
```

## Decisions and why

### C# on .NET, late-bound COM
Outlook is a COM server. A .NET process can call it directly, in-process, with no intermediate shell. An earlier approach that generated PowerShell text per call suffered the whole class of bugs that come from building code in strings (quoting, bare `null`, culture-sensitive dates) and paid one to two seconds of process start-up per call. Calling COM directly removes both.

COM is used late-bound (`dynamic`). This avoids shipping Office interop assemblies and works with any installed Outlook version. The cost is that the code cannot be trimmed, which is why the release is a 39 MB self-contained file rather than a few megabytes.

### One STA thread
Outlook's object model is single-threaded-apartment COM. `Com/Sta.cs` runs every COM call on one dedicated STA thread through a queue. COM objects never cross apartments and calls are serialised without locks.

A COM call cannot be aborted. If Outlook is blocked (a modal dialog, a sign-in prompt), the caller is released after `OUTLOOK_MCP_TIMEOUT_SECONDS` with an actionable message. The abandoned job is flagged so that, if the thread frees up later, it is **dropped rather than executed**. While the thread is still stuck, new calls fail fast. This matters most for sends.

### Session and reconnect
`Com/Ol.cs` creates `Outlook.Application` lazily and recovers once from the RPC errors raised when Outlook is closed or restarted (`RPC_E_DISCONNECTED` and friends). Everything else in that file is small helpers: safe property reads, folder and account resolution, recipient SMTP resolution, and id handling.

### Compact ids
A MAPI EntryID is about 48 hex characters, and every item in a store starts with the same 40 (flags plus the store's provider UID). The server hands out `"<6-hex store alias>:<entry id without that prefix>"`, for example `7ecb73:E4362200`. The alias is the first three bytes of a SHA-1 of the store id. If an id ever lacks the expected prefix it is emitted whole after a `~`, so the scheme degrades gracefully. Ids are stable across restarts.

### Sending as a specific account
For POP and IMAP accounts Outlook either ignores `MailItem.SendUsingAccount` or fails with `RPC_E_SERVERFAULT`, even for the default account. The mechanism Outlook does honour is the store that holds the item. With `account=`, the message is created in that account's own Drafts folder and checked to be in the right store before sending. Verified end to end: messages created this way arrive from the chosen account and leave through that account's own SMTP route.

### Tool registry
Tools are plain static methods with `[ToolMeta(name, toolset, access)]` and `[Description]` attributes. `Registry` discovers them by reflection and applies policy, so the server, the generated docs and the tests all see the same list. `access` (`Read`, `Write`, `Send`, `Delete`) drives:

- the MCP annotations (`readOnlyHint`, `destructiveHint`, `openWorldHint`),
- `--read-only` and `--disable-send`.

Unit tests enforce naming, that every tool and parameter has a description, and a token budget for the whole tool surface.

### Output is a product
Results are text, shaped for a language model rather than for humans or machines:

- one line per item, fixed field order, flags as short letters;
- bodies pass through `Fmt.CleanBody`: angle-bracket links and `mailto:` noise removed, gateway banners stripped, the quoted reply chain cut, double-spaced text collapsed, trailing legal footers dropped;
- long bodies return an exact `body_offset` to continue from;
- no JSON envelope and no repeated labels.

### Errors
Expected failures throw `ToolError` with a message written for the model ("Drafts cannot be flagged for follow-up…"). A call-tool filter turns every exception into a normal `isError` result, so the protocol stays clean and no stack traces leak. In debug mode the stack is logged to stderr and the file log.

## Layout

```
src/OutlookMcp/
  Program.cs        entry point and host wiring
  Cli.cs            doctor, install, install-skill, tools
  Settings.cs       flags and environment, policy decisions
  Registry.cs       tool discovery and policy
  Debug.cs          debug log, argument redaction
  Fmt.cs            compact formatting and body cleaning
  Com/Ol.cs         Outlook session, ids, folders, recipients
  Com/Sta.cs        STA thread, watchdog
  Tools/            MailTools, CalendarTools, ContactTools, TaskTools, SystemTools
skills/outlook/     agent skill (embedded in the executable)
tests/              unit and installer tests
scripts/            live integration suite and a stdio probe
```
