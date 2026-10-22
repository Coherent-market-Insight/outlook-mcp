# Troubleshooting

Start with `outlook-mcp doctor`. It checks the OS, whether Outlook's COM class is registered, connects, lists accounts, and times an inbox read.

## "Outlook is not installed or not registered for COM"

You are probably on the **new Outlook for Windows**, which has no COM interface, or Outlook is not installed. Use classic Outlook (Microsoft 365 Apps, or Outlook 2016/2019/2021). In the new Outlook, switch back using the toggle in the top-right corner of the window.

## "Outlook did not answer within 60s"

Outlook is blocked, usually by a dialog box: a security prompt, a sign-in or password request, a "send/receive" window, or an "offline" notice. Bring Outlook to the front, close the dialog, and retry. Calls queued behind a blocked one are dropped on purpose so nothing stale runs later. Raise the wait with `OUTLOOK_MCP_TIMEOUT_SECONDS` if a very large mailbox needs longer.

## A security prompt ("a program is trying to access email addresses")

Outlook shows this when it does not detect up-to-date antivirus. Fix the antivirus state (Windows Security counts), or allow access once from the prompt.

## "Could not create the message in <account>'s mailbox"

The account's store could not be opened for writing, so sending was refused rather than risking the wrong sender. Check the account in Outlook (File, Account Settings), make sure its data file is not read-only, and retry.

## Sends go from the wrong account

Use `account` with the account's SMTP address from `status`. Without it, Outlook uses the default account. Replies go out from the mailbox the original message arrived in.

## The client shows the server as disconnected

Run the executable by hand in a terminal. It should wait silently on stdin. Then run it with `--debug` and read stderr, or add `--log-file C:\temp\outlook-mcp.log` to the client entry and read the file. Common causes: the path in the config no longer exists (re-run `install`), or the client was not restarted.

## Search returns nothing for text you can see in Outlook

Full-text search uses Outlook's own filters on subject, sender, recipients and body. Mail that has not finished indexing or downloading will not match. Try `subject` or `from`, or widen with `since`.

## Dates look wrong

Dates are local time in ISO 8601 (`2026-10-30T14:00`). A date without a time means midnight. All-day events take dates, and `end` is the last day, inclusive.

## Free/busy for other people says "no data"

Other people's availability needs an Exchange or Microsoft 365 account. POP and IMAP accounts cannot supply it.

## Reading the debug log

Each call produces `> tool {args}` and `< tool ok|error <ms> <chars> chars (~tokens)`. Text arguments appear as `str(<length>)`, never their content, so logs are safe to share.

## Still stuck

Open an issue with the output of `outlook-mcp version` and `outlook-mcp doctor`. Remove addresses you do not want to share.
