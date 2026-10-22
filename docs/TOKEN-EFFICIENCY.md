# Token efficiency

Every design choice here trades a little convenience for fewer tokens. The tool-definition sizes below are measured on this build; per-call figures are estimates. One token is taken as about four characters.

## Standing cost: tool definitions

A connected MCP server costs tokens on every request through its tool list, so this is the number that matters most.

| Configuration | Tools | Definitions |
|---|---|---|
| Everything | 31 | ~6.6K tokens |
| `--read-only` | 13 | ~2.4K tokens |
| `--toolsets mail` | 15 | ~3.1K tokens |
| `--toolsets mail --read-only` | 5 | ~1.1K tokens |
| `--toolsets calendar` | 9 | ~2.2K tokens |

Reduce it by exposing only what a job needs. A mail-triage agent rarely needs calendar, contacts or tasks, and a reporting agent never needs the write tools. Clients that load tool definitions on demand reduce the standing cost further.

A unit test fails the build if the full surface grows past a fixed budget, so descriptions cannot creep.

## Per-call cost

| Technique | Effect |
|---|---|
| **Short ids** | `7ecb73:E4362200` (15 characters) instead of a 48-character EntryID. Every message line and every follow-up call carries one. |
| **One line per item** | `id date sender \| subject [flags]`, roughly 25 to 35 tokens. No JSON envelope, no repeated field names. |
| **Previews are opt-in** | `preview` defaults to 0. Ask for 80 to 120 characters only when triaging by content. |
| **Filter server-side** | `query`, `from`, `to`, `subject`, `unread`, `attachments`, `flagged`, `category`, `since`, `until` run inside Outlook, so non-matching mail never reaches the model. |
| **Cleaned bodies** | `mail_get` removes `<https://...>` and `<mailto:...>` noise (tracking links alone can be hundreds of tokens), gateway banners, the quoted reply chain, trailing legal footers, and double spacing. |
| **Capped bodies** | Default 3000 characters, with an exact `body_offset` to continue. Nothing is silently truncated. |
| **Batches** | `mail_update`, `mail_move`, `mail_delete` take many ids in one call. |
| **Threads** | `mail_thread` returns trimmed bodies for a whole conversation in one call instead of one `mail_get` per message. |
| **Attachments by path** | Contents are never returned inline; `mail_attachments save=true` returns file paths. |
| **Lean writes** | Write tools answer with one confirmation line (`sent → a@b.com \| subject`). |

## Typical session (estimates)

| Step | Calls | Approx. tokens returned |
|---|---|---|
| List 15 unread messages | 1 | ~400 |
| Read 3 of them | 3 | ~900 to 1,500 |
| Mark 12 as read, move 3 | 2 | ~40 |

## Tuning tips

- Prefer `since` to paging. `mail_list since=2026-10-20 unread=true` is one small call.
- Search with `query` before listing a whole folder.
- Keep `limit` modest (15 to 25). Use `offset` from the `more:` hint to continue.
- Lower `OUTLOOK_MCP_BODY_MAX_CHARS` for agents that only need the gist.
- Use `mail_thread` for context, not repeated `mail_get`.
