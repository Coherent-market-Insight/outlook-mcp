# Tool reference

Generated from the code with `outlook-mcp tools --markdown`. **Access** drives the MCP annotations and what `--read-only` (keeps `read` only) and `--disable-send` (removes `send`) hide.

Ids come from list and search results. Dates are local-time ISO 8601. Tools that change data say so in their description; sending tools never send unless the description says they send.

### calendar

| Tool | Access | What it does |
|---|---|---|
| `calendar_event_create` | write | Create an event or meeting. With attendees it is saved unsent unless send_invites=true. Supports all-day events, reminders, show-as and simple recurrence. |
| `calendar_event_delete` | delete | Delete an event, or one occurrence of a recurring event. With cancel_attendees=true a meeting you organise is cancelled and attendees are notified. |
| `calendar_event_get` | read | Full details of one event: time, location, organizer, attendees with their responses, recurrence, reminder, body. |
| `calendar_event_update` | write | Change an event. Only the fields you pass change. For a recurring event pass occurrence_start to change a single occurrence. |
| `calendar_events` | read | Events in a date range, in time order, recurring meetings expanded. One line each: id, time, subject, location, flags R=recurring M=meeting. For one occurrence of a recurring event pass its start as occurrence_start to other tools. |
| `calendar_free_busy` | read | Find free slots in working hours. Uses your calendar, plus other people's free/busy when attendees are given (Exchange/M365 only; others report no data). |
| `calendar_list` | read | List calendars across all accounts. Use a returned name as the `calendar` argument of other calendar tools. |
| `calendar_respond` | send | Respond to a meeting invitation: accept, tentative or decline. Pass the invitation's message id from the inbox, or an event id. |

### contacts

| Tool | Access | What it does |
|---|---|---|
| `contact_delete` | delete | Delete contacts (moved to Deleted Items). |
| `contact_get` | read | All stored fields of one contact. |
| `contact_save` | write | Create a contact, or update one when id is given (only the fields you pass change). |
| `contacts_search` | read | Search saved contacts by name, email or company. One line each: id, name, company, title, email, phone. |
| `recipient_resolve` | read | Resolve names or partial addresses to email addresses using saved contacts and, on Exchange, the company address list. Use before emailing someone you only know by name. |

### mail

| Tool | Access | What it does |
|---|---|---|
| `mail_attachments` | write | List a message's attachments, or save them to disk (save=true) and get the file paths. Contents are never returned inline. |
| `mail_delete` | delete | Delete messages: moves them to Deleted Items, or removes them for good with permanent=true. |
| `mail_draft_save` | write | Create a draft, or update one when id is given (only the fields you pass change). Returns the draft id. Nothing is sent. |
| `mail_draft_send` | send | Send an existing draft. |
| `mail_folder_create` | write | Create a mail folder. |
| `mail_folders` | read | Folder tree with unread/total counts for an account (or every account). Use it to find folder paths. |
| `mail_forward` | send | Forward a message with its attachments. Saves a draft unless send=true. |
| `mail_get` | read | Read one message: headers, attachment list and a cleaned body (links stripped, quoted reply chain cut). Does not mark it read. |
| `mail_list` | read | List or search mail in a folder, newest first. One line per message: id, date, sender (recipient in Sent/Drafts), subject, flags U=unread !=high @=attachment F=flagged. Use ids with mail_get. account="all" merges the same folder across every mailbox. |
| `mail_move` | write | Move messages to another folder (for example archive or a client folder). |
| `mail_reply` | send | Reply to a message. Saves a draft unless send=true. Your text goes above the quoted original and the original's formatting is preserved. |
| `mail_send` | send | Send a new email immediately. To review first, use mail_draft_save. Plain text unless html=true. No signature is added. |
| `mail_thread` | read | Show the conversation a message belongs to (Inbox, Sent, Drafts), oldest first, with trimmed bodies. |
| `mail_update` | write | Change flags on one or many messages: read state, follow-up flag, categories, importance. |

### system

| Tool | Access | What it does |
|---|---|---|
| `status` | read | Outlook version, signed-in accounts (with SMTP address and type), default account and store, and this server's active policy. Call once to learn which accounts exist. |

### tasks

| Tool | Access | What it does |
|---|---|---|
| `task_delete` | delete | Delete tasks (moved to Deleted Items). |
| `task_save` | write | Create a task, or update one when id is given (only the fields you pass change). Set status=completed to finish a task. |
| `tasks_list` | read | List tasks, soonest due first. One line each: id, due date, subject, status, flags !=high. |


## Toolset names

`mail`, `calendar`, `contacts`, `tasks`; `status` is always available. Use `--toolsets` to limit them.
