using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
using OutlookMcp.Com;

namespace OutlookMcp.Tools;

internal static partial class MailTools
{
    private const string Dasl = "urn:schemas:httpmail:";

    // ------------------------------------------------------------------ list / search

    [ToolMeta("mail_list", "mail", Access.Read)]
    [Description("List or search mail in a folder, newest first. One line per message: id, date, sender (recipient in Sent/Drafts), subject, flags U=unread !=high @=attachment F=flagged. Use ids with mail_get. account=\"all\" merges the same folder across every mailbox.")]
    public static Task<string> MailList(
        [Description("inbox (default), sent, drafts, deleted, junk, outbox, a path like Inbox/Clients, or a folder id")] string? folder = null,
        [Description("Account SMTP address or name, or \"all\" for every mailbox; default account otherwise")] string? account = null,
        [Description("Free text matched in subject, sender, recipients and body")] string? query = null,
        [Description("Sender name or address contains")] string? from = null,
        [Description("Recipient contains")] string? to = null,
        [Description("Subject contains")] string? subject = null,
        [Description("true: only unread; false: only read")] bool? unread = null,
        [Description("true: only with attachments")] bool? attachments = null,
        [Description("true: only flagged for follow-up")] bool? flagged = null,
        [Description("Category name")] string? category = null,
        [Description("Newer than, ISO date or datetime")] string? since = null,
        [Description("Older than, ISO date or datetime")] string? until = null,
        [Description("Preview characters per message, 0 = none (cheapest)")] int preview = 0,
        [Description("Max rows (1-100)")] int limit = 15,
        [Description("Rows to skip, for paging")] int offset = 0)
        => Sta.Run<string>(() =>
        {
            limit = Math.Clamp(limit, 1, 100);
            offset = Math.Max(0, offset);
            var filter = new ListFilter(query, from, to, subject, unread, attachments, flagged, category,
                since is null ? null : Ol.ParseDate(since, "since"), until is null ? null : Ol.ParseDate(until, "until"), preview);

            var all = string.Equals(account, "all", StringComparison.OrdinalIgnoreCase);
            var sources = new List<string?>();
            if (all)
            {
                var seen = new HashSet<string>();
                foreach (dynamic a in Ol.Ns.Accounts)
                    if (seen.Add((string)a.DeliveryStore.StoreID)) sources.Add((string)a.SmtpAddress);
            }
            else sources.Add(account);

            var rows = new List<Row>();
            var names = new List<string>();
            var total = 0;
            var dated = filter.Since is not null || filter.Until is not null || filter.Flagged == true;
            var cap = offset + limit + 1;
            foreach (var src in sources)
            {
                var r = ListCore(folder, src, filter, cap);
                rows.AddRange(r.Rows);
                total += r.Total;
                names.Add(all ? $"{r.Name}[{src}]" : r.Name);
            }
            var ordered = rows.OrderByDescending(r => r.When).ToList();
            var page = ordered.Skip(offset).Take(limit).ToList();
            var more = ordered.Count > offset + limit;
            var head = $"{(all ? string.Join(" + ", names) : names[0])}: {page.Count} shown" + (dated ? "" : $", {total} match") +
                       (more ? $", more: offset={offset + page.Count}" : "");
            return page.Count == 0 ? head + " (none)" : head + "\n" + string.Join("\n", page.Select(r => r.Line));
        });

    private sealed record ListFilter(string? Query, string? From, string? To, string? Subject, bool? Unread, bool? Attachments,
        bool? Flagged, string? Category, DateTime? Since, DateTime? Until, int Preview);

    private sealed record Row(DateTime When, string Line);

    private static (List<Row> Rows, string Name, int Total) ListCore(string? folder, string? account, ListFilter f, int cap)
    {
        var fld = Ol.ResolveFolder(folder, account, Ol.Inbox);
        if ((int)fld.DefaultItemType != 0) throw new ToolError($"'{(string)fld.Name}' is not a mail folder.");

        var conds = new List<string>();
        static string Esc(string s) => s.Replace("'", "''");
        if (!string.IsNullOrWhiteSpace(f.Query))
        {
            var q = Esc(f.Query!);
            conds.Add($"(\"{Dasl}subject\" LIKE '%{q}%' OR \"{Dasl}fromname\" LIKE '%{q}%' OR \"{Dasl}fromemail\" LIKE '%{q}%' OR \"{Dasl}displayto\" LIKE '%{q}%' OR \"{Dasl}textdescription\" LIKE '%{q}%')");
        }
        if (!string.IsNullOrWhiteSpace(f.From)) conds.Add($"(\"{Dasl}fromname\" LIKE '%{Esc(f.From!)}%' OR \"{Dasl}fromemail\" LIKE '%{Esc(f.From!)}%')");
        if (!string.IsNullOrWhiteSpace(f.To)) conds.Add($"(\"{Dasl}displayto\" LIKE '%{Esc(f.To!)}%' OR \"{Dasl}displaycc\" LIKE '%{Esc(f.To!)}%')");
        if (!string.IsNullOrWhiteSpace(f.Subject)) conds.Add($"\"{Dasl}subject\" LIKE '%{Esc(f.Subject!)}%'");
        if (f.Unread is not null) conds.Add($"\"{Dasl}read\" = {(f.Unread.Value ? 0 : 1)}");
        if (f.Attachments == true) conds.Add($"\"{Dasl}hasattachment\" = 1");
        if (!string.IsNullOrWhiteSpace(f.Category)) conds.Add($"\"urn:schemas-microsoft-com:office:office#Keywords\" LIKE '%{Esc(f.Category!)}%'");

        var kind = FolderKind(fld);
        var dateProp = kind == Ol.Sent ? "SentOn" : kind == Ol.Drafts ? "LastModificationTime" : "ReceivedTime";
        var showTo = kind is Ol.Sent or Ol.Drafts or Ol.Outbox;

        dynamic items = fld.Items;
        if (conds.Count > 0) items = items.Restrict("@SQL=" + string.Join(" AND ", conds));
        items.Sort($"[{dateProp}]", true);
        int total = items.Count;

        var rows = new List<Row>();
        dynamic? it = items.GetFirst();
        while (it is not null && rows.Count < cap)
        {
            DateTime when = Ol.Safe(() => (DateTime)(dateProp == "SentOn" ? it.SentOn : dateProp == "LastModificationTime" ? it.LastModificationTime : it.ReceivedTime), DateTime.MinValue);
            var stop = f.Since is not null && when < f.Since;
            var skip = f.Until is not null && when > f.Until;
            var unflagged = f.Flagged == true && !Ol.Safe(() => (int)it.FlagStatus == 2, false);
            if (!stop && !skip && !unflagged) rows.Add(new Row(when, MailLine(it, when, showTo, f.Preview)));
            Ol.Release(it);
            if (stop) break;
            it = items.GetNext();
        }
        return (rows, (string)fld.Name, f.Flagged == true ? -1 : total);
    }

    private static string MailLine(dynamic mi, DateTime when, bool showTo, int preview)
    {
        var id = Ol.Id(mi);
        var who = showTo
            ? "→ " + Fmt.OneLine(Ol.S(() => mi.To), 50)
            : Fmt.OneLine(Ol.S(() => mi.SenderName), 40);
        var line = $"{id} {Fmt.Date(when)} {who} | {Fmt.OneLine(Ol.S(() => mi.Subject), 110)}" +
                   Fmt.Flags(
                       unread: Ol.Safe(() => (bool)mi.UnRead, false),
                       high: Ol.Safe(() => (int)mi.Importance == 2, false),
                       attach: Ol.Safe(() => (int)mi.Attachments.Count > 0, false),
                       flagged: Ol.Safe(() => (int)mi.FlagStatus == 2, false));
        if (preview > 0)
            line += "\n  " + Fmt.OneLine(Fmt.CleanBody(Ol.S(() => mi.Body), false, false), Math.Min(preview, 400));
        return line;
    }

    private static int FolderKind(dynamic folder)
    {
        try
        {
            string eid = folder.EntryID;
            foreach (var kind in new[] { Ol.Sent, Ol.Drafts, Ol.Outbox, Ol.Deleted, Ol.Junk })
            {
                dynamic d = folder.Store.GetDefaultFolder(kind);
                if ((string)d.EntryID == eid) return kind;
            }
        }
        catch { }
        return Ol.Inbox;
    }

    // ------------------------------------------------------------------ read

    [ToolMeta("mail_get", "mail", Access.Read)]
    [Description("Read one message: headers, attachment list and a cleaned body (links stripped, quoted reply chain cut). Does not mark it read.")]
    public static Task<string> MailGet(
        [Description("Message id from mail_list")] string id,
        [Description("Max body characters; default 3000")] int? max_chars = null,
        [Description("Body start offset, for long messages")] int body_offset = 0,
        [Description("true: keep URLs and mailto links")] bool links = false,
        [Description("true: keep the quoted reply chain")] bool quoted = false,
        [Description("true: include To/Cc/Bcc with addresses and message headers")] bool headers = false)
        => Sta.Run<string>(() =>
        {
            dynamic mi = Ol.GetItem(id);
            var sb = new StringBuilder();
            var cls = Ol.Safe(() => (int)mi.Class, 43);
            sb.AppendLine($"id: {id}");
            sb.AppendLine($"subject: {Ol.S(() => mi.Subject)}");
            var sender = Ol.S(() => mi.SenderName);
            sb.AppendLine(sender.Length > 0 ? $"from: {sender} <{SenderSmtp(mi)}>" : "from: (draft)");
            sb.AppendLine($"date: {Fmt.Iso(Ol.Safe(() => (DateTime)mi.ReceivedTime, Ol.Safe(() => (DateTime)mi.CreationTime, DateTime.MinValue)))}");
            AppendRecipients(sb, mi, headers);
            var flags = Fmt.Flags(
                unread: Ol.Safe(() => (bool)mi.UnRead, false),
                high: Ol.Safe(() => (int)mi.Importance == 2, false),
                flagged: Ol.Safe(() => (int)mi.FlagStatus == 2, false));
            if (flags.Length > 0) sb.AppendLine($"flags:{flags}");
            var cats = Ol.S(() => mi.Categories);
            if (cats.Length > 0) sb.AppendLine($"categories: {cats}");
            if (cls == 53) sb.AppendLine("type: meeting request (use calendar_respond)");
            AppendAttachments(sb, mi);
            if (headers)
            {
                var raw = Ol.S(() => mi.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"));
                if (raw.Length > 0) sb.AppendLine("headers:\n" + Fmt.OneLine(raw, 1500));
            }

            var body = Ol.S(() => mi.Body);
            if (body.Trim().Length == 0) body = Fmt.StripHtml(Ol.S(() => mi.HTMLBody));
            var clean = Fmt.CleanBody(body, links, quoted);
            sb.AppendLine("--");
            sb.Append(Fmt.Window(clean, Math.Max(0, body_offset), max_chars ?? Current.Settings.BodyMaxChars, "body_offset"));
            return sb.ToString();
        });

    [ToolMeta("mail_thread", "mail", Access.Read)]
    [Description("Show the conversation a message belongs to (Inbox, Sent, Drafts), oldest first, with trimmed bodies.")]
    public static Task<string> MailThread(
        [Description("Any message id in the conversation")] string id,
        [Description("Max messages")] int limit = 15,
        [Description("Max body characters per message")] int max_chars = 600)
        => Sta.Run<string>(() =>
        {
            dynamic seed = Ol.GetItem(id);
            string topic = seed.ConversationTopic;
            dynamic store = seed.Parent.Store;
            var rows = new List<(DateTime When, string Text)>();
            foreach (var kind in new[] { Ol.Inbox, Ol.Sent, Ol.Drafts })
            {
                dynamic fld;
                try { fld = store.GetDefaultFolder(kind); } catch { continue; }
                dynamic items = fld.Items.Restrict($"[ConversationTopic] = '{topic.Replace("'", "''")}'");
                dynamic? it = items.GetFirst();
                while (it is not null)
                {
                    var when = Ol.Safe(() => (DateTime)(kind == Ol.Inbox ? it.ReceivedTime : kind == Ol.Sent ? it.SentOn : it.LastModificationTime), DateTime.MinValue);
                    var body = Fmt.CleanBody(Ol.S(() => it.Body), false, false);
                    var head = $"{Ol.Id(it)} {Fmt.Date(when)} {(kind == Ol.Inbox ? Ol.S(() => it.SenderName) : "me → " + Fmt.OneLine(Ol.S(() => it.To), 40))}" +
                               Fmt.Flags(unread: Ol.Safe(() => (bool)it.UnRead, false), attach: Ol.Safe(() => (int)it.Attachments.Count > 0, false));
                    rows.Add((when, head + "\n" + Fmt.OneLine(body, max_chars)));
                    Ol.Release(it);
                    it = items.GetNext();
                }
            }
            if (rows.Count == 0) return "No conversation found.";
            var ordered = rows.OrderBy(r => r.When).ToList();
            var take = ordered.Count > limit ? ordered.Skip(ordered.Count - limit).ToList() : ordered;
            return $"{topic} ({ordered.Count} messages{(ordered.Count > take.Count ? $", showing last {take.Count}" : "")})\n" +
                   string.Join("\n\n", take.Select(r => r.Text));
        });

    // ------------------------------------------------------------------ compose

    [ToolMeta("mail_send", "mail", Access.Send)]
    [Description("Send a new email immediately. To review first, use mail_draft_save. Plain text unless html=true. No signature is added.")]
    public static Task<string> MailSend(
        [Description("To addresses")] string[] to,
        [Description("Subject")] string subject,
        [Description("Body")] string body,
        [Description("Cc addresses")] string[]? cc = null,
        [Description("Bcc addresses")] string[]? bcc = null,
        [Description("true: body is HTML")] bool html = false,
        [Description("Absolute file paths to attach")] string[]? attachments = null,
        [Description("Send from this account (SMTP address or name); default account otherwise")] string? account = null,
        [Description("low, normal or high")] string? importance = null,
        [Description("Delay delivery until this ISO datetime")] string? defer_until = null)
        => Sta.Run<string>(() =>
        {
            var mi = Compose(to, cc, bcc, subject, body, html, attachments, account, importance, defer_until);
            var summary = RecipientSummary(mi);
            Guard(mi);
            mi.Send();
            return $"sent → {summary} | {subject}";
        });

    [ToolMeta("mail_draft_save", "mail", Access.Write)]
    [Description("Create a draft, or update one when id is given (only the fields you pass change). Returns the draft id. Nothing is sent.")]
    public static Task<string> MailDraftSave(
        [Description("Existing draft id to update; omit to create")] string? id = null,
        [Description("To addresses")] string[]? to = null,
        [Description("Subject")] string? subject = null,
        [Description("Body")] string? body = null,
        [Description("Cc addresses")] string[]? cc = null,
        [Description("Bcc addresses")] string[]? bcc = null,
        [Description("true: body is HTML")] bool html = false,
        [Description("Absolute file paths to attach (added)")] string[]? attachments = null,
        [Description("Account for a new draft")] string? account = null,
        [Description("low, normal or high")] string? importance = null)
        => Sta.Run<string>(() =>
        {
            dynamic mi;
            if (id is null)
            {
                if (to is null || to.Length == 0 || subject is null || body is null)
                    throw new ToolError("Creating a draft needs to, subject and body.");
                mi = Compose(to, cc, bcc, subject, body, html, attachments, account, importance, null);
            }
            else
            {
                mi = Ol.GetItem(id);
                if (!Ol.Safe(() => (bool)mi.Sent == false, false)) throw new ToolError("That message was already sent; it is not a draft.");
                if (to is not null) mi.To = string.Join(";", to);
                if (cc is not null) mi.CC = string.Join(";", cc);
                if (bcc is not null) mi.BCC = string.Join(";", bcc);
                if (subject is not null) mi.Subject = subject;
                if (body is not null) { if (html) mi.HTMLBody = body; else mi.Body = body; }
                if (importance is not null) mi.Importance = Importance(importance);
                AddAttachments(mi, attachments);
            }
            mi.Save();
            return $"draft saved: {Ol.Id(mi)} → {RecipientSummary(mi)} | {Ol.S(() => mi.Subject)}";
        });

    [ToolMeta("mail_draft_send", "mail", Access.Send)]
    [Description("Send an existing draft.")]
    public static Task<string> MailDraftSend([Description("Draft id")] string id)
        => Sta.Run<string>(() =>
        {
            dynamic mi = Ol.GetItem(id);
            if (!Ol.Safe(() => (bool)mi.Sent == false, false)) throw new ToolError("That message was already sent.");
            var summary = RecipientSummary(mi);
            Guard(mi);
            mi.Send();
            return $"sent → {summary} | {Ol.S(() => mi.Subject)}";
        });

    [ToolMeta("mail_reply", "mail", Access.Send)]
    [Description("Reply to a message. Saves a draft unless send=true. Your text goes above the quoted original and the original's formatting is preserved.")]
    public static Task<string> MailReply(
        [Description("Message id")] string id,
        [Description("Your reply text")] string body,
        [Description("true: reply to all")] bool all = false,
        [Description("true: send now; default saves a draft")] bool send = false,
        [Description("true: body is HTML")] bool html = false,
        [Description("Extra Cc addresses")] string[]? cc = null,
        [Description("Absolute file paths to attach")] string[]? attachments = null)
        => Sta.Run<string>(() =>
        {
            dynamic src = Ol.GetItem(id);
            dynamic r = all ? src.ReplyAll() : src.Reply();
            PrependBody(r, body, html);
            if (cc is not null && cc.Length > 0) r.CC = string.Join(";", new[] { Ol.S(() => r.CC) }.Concat(cc).Where(s => s.Length > 0));
            AddAttachments(r, attachments);
            return FinishComposed(r, send, "reply");
        });

    [ToolMeta("mail_forward", "mail", Access.Send)]
    [Description("Forward a message with its attachments. Saves a draft unless send=true.")]
    public static Task<string> MailForward(
        [Description("Message id")] string id,
        [Description("To addresses")] string[] to,
        [Description("Optional note above the forwarded message")] string? body = null,
        [Description("true: send now; default saves a draft")] bool send = false,
        [Description("true: note is HTML")] bool html = false,
        [Description("Extra Cc addresses")] string[]? cc = null)
        => Sta.Run<string>(() =>
        {
            dynamic src = Ol.GetItem(id);
            dynamic f = src.Forward();
            f.To = string.Join(";", to);
            if (cc is not null && cc.Length > 0) f.CC = string.Join(";", cc);
            if (!string.IsNullOrEmpty(body)) PrependBody(f, body!, html);
            return FinishComposed(f, send, "forward");
        });

    // ------------------------------------------------------------------ organise

    [ToolMeta("mail_update", "mail", Access.Write)]
    [Description("Change flags on one or many messages: read state, follow-up flag, categories, importance.")]
    public static Task<string> MailUpdate(
        [Description("Message ids")] string[] ids,
        [Description("true: mark read; false: mark unread")] bool? read = null,
        [Description("none, flagged or complete")] string? flag = null,
        [Description("Replace categories with these")] string[]? categories = null,
        [Description("Add these categories")] string[]? add_categories = null,
        [Description("Remove these categories")] string[]? remove_categories = null,
        [Description("low, normal or high")] string? importance = null)
        => Sta.Run<string>(() =>
        {
            var done = 0; var failed = new List<string>();
            foreach (var id in ids)
            {
                try
                {
                    dynamic mi = Ol.GetItem(id);
                    if (read is not null) mi.UnRead = !read.Value;
                    if (flag is not null)
                    {
                        switch (flag.ToLowerInvariant())
                        {
                            case "flagged": RequireSent(mi); mi.MarkAsTask(0); mi.FlagStatus = 2; break;
                            case "complete": RequireSent(mi); mi.FlagStatus = 1; break;
                            case "none": if (!Ol.Safe(() => (bool)mi.Sent == false, false)) mi.ClearTaskFlag(); break;
                            default: throw new ToolError("flag must be none, flagged or complete.");
                        }
                    }
                    if (categories is not null || add_categories is not null || remove_categories is not null)
                    {
                        var set = (categories is not null
                            ? categories.ToList()
                            : Ol.S(() => mi.Categories).Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList());
                        if (add_categories is not null) foreach (var c in add_categories) if (!set.Contains(c, StringComparer.OrdinalIgnoreCase)) set.Add(c);
                        if (remove_categories is not null) set.RemoveAll(c => remove_categories.Contains(c, StringComparer.OrdinalIgnoreCase));
                        mi.Categories = string.Join(", ", set);
                    }
                    if (importance is not null) mi.Importance = Importance(importance);
                    mi.Save();
                    done++;
                    Ol.Release(mi);
                }
                catch (Exception ex) { failed.Add($"{Ol.Trunc(id, 14)}…: {ex.Message}"); }
            }
            return $"updated {done}/{ids.Length}" + (failed.Count > 0 ? "\nfailed: " + string.Join("; ", failed) : "");
        });

    [ToolMeta("mail_move", "mail", Access.Write)]
    [Description("Move messages to another folder (for example archive or a client folder).")]
    public static Task<string> MailMove(
        [Description("Message ids")] string[] ids,
        [Description("Destination folder: sent, deleted, junk, a path like Inbox/Clients, or a folder id")] string folder,
        [Description("Account for default-folder names")] string? account = null)
        => Sta.Run<string>(() =>
        {
            dynamic dest = Ol.ResolveFolder(folder, account, Ol.Inbox);
            var done = 0;
            foreach (var id in ids)
            {
                dynamic mi = Ol.GetItem(id);
                mi.Move(dest);
                done++;
            }
            return $"moved {done}/{ids.Length} → {(string)dest.Name}";
        });

    [ToolMeta("mail_delete", "mail", Access.Delete)]
    [Description("Delete messages: moves them to Deleted Items, or removes them for good with permanent=true.")]
    public static Task<string> MailDelete(
        [Description("Message ids")] string[] ids,
        [Description("true: delete permanently (cannot be undone)")] bool permanent = false)
        => Sta.Run<string>(() =>
        {
            var done = 0;
            foreach (var id in ids)
            {
                dynamic mi = Ol.GetItem(id);
                if (permanent)
                {
                    dynamic moved = mi.Move(Ol.DefaultFolder(Ol.Deleted, null));
                    (moved ?? mi).Delete();
                }
                else mi.Delete();
                done++;
            }
            return $"{(permanent ? "permanently deleted" : "deleted")} {done}/{ids.Length}";
        });

    // ------------------------------------------------------------------ attachments

    [ToolMeta("mail_attachments", "mail", Access.Write)]
    [Description("List a message's attachments, or save them to disk (save=true) and get the file paths. Contents are never returned inline.")]
    public static Task<string> MailAttachments(
        [Description("Message id")] string id,
        [Description("true: save files to disk")] bool save = false,
        [Description("Folder to save into; default a temp folder")] string? directory = null,
        [Description("Only save these file names")] string[]? names = null)
        => Sta.Run<string>(() =>
        {
            dynamic mi = Ol.GetItem(id);
            var sb = new StringBuilder();
            int count = mi.Attachments.Count;
            if (count == 0) return "no attachments";
            var dir = directory ?? Path.Combine(Path.GetTempPath(), "outlook-mcp", Ol.Trunc(id.Split(':').Last(), 12));
            if (save) Directory.CreateDirectory(dir);
            for (var i = 1; i <= count; i++)
            {
                dynamic a = mi.Attachments[i];
                string name = a.FileName;
                var inline = Ol.S(() => a.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")).Length > 0;
                var size = Ol.Safe(() => (long)a.Size, 0L);
                if (save && (names is null || names.Contains(name, StringComparer.OrdinalIgnoreCase)))
                {
                    var safe = string.Concat(name.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                    var path = Path.Combine(dir, safe);
                    a.SaveAsFile(path);
                    sb.AppendLine($"{i}. {name} ({Fmt.Size(size)}) → {path}");
                }
                else sb.AppendLine($"{i}. {name} ({Fmt.Size(size)}){(inline ? " inline" : "")}");
                Ol.Release(a);
            }
            return sb.ToString().TrimEnd();
        });

    // ------------------------------------------------------------------ folders

    [ToolMeta("mail_folders", "mail", Access.Read)]
    [Description("Folder tree with unread/total counts for an account (or every account). Use it to find folder paths.")]
    public static Task<string> MailFolders(
        [Description("Account SMTP address or name; omit for all accounts")] string? account = null,
        [Description("Tree depth, 1-4")] int depth = 2)
        => Sta.Run<string>(() =>
        {
            depth = Math.Clamp(depth, 1, 4);
            var sb = new StringBuilder();
            void Walk(dynamic folder, int level)
            {
                if ((int)folder.DefaultItemType == 0 || level == 0)
                {
                    int total = Ol.Safe(() => (int)folder.Items.Count, 0);
                    int unread = Ol.Safe(() => (int)folder.UnReadItemCount, 0);
                    if (level > 0) sb.AppendLine($"{new string(' ', (level - 1) * 2)}{(string)folder.Name}  {unread}/{total}");
                }
                if (level >= depth) return;
                foreach (dynamic child in folder.Folders)
                    if ((int)child.DefaultItemType == 0) Walk(child, level + 1);
            }

            if (account is not null)
            {
                dynamic acct = Ol.FindAccount(account)!;
                sb.AppendLine($"# {(string)acct.SmtpAddress}");
                Walk(acct.DeliveryStore.GetRootFolder(), 0);
                foreach (dynamic c in acct.DeliveryStore.GetRootFolder().Folders) if ((int)c.DefaultItemType == 0) Walk(c, 1);
            }
            else
            {
                foreach (dynamic st in Ol.Ns.Stores)
                {
                    sb.AppendLine($"# {(string)st.DisplayName}");
                    foreach (dynamic c in st.GetRootFolder().Folders) if ((int)c.DefaultItemType == 0) Walk(c, 1);
                }
            }
            return sb.ToString().TrimEnd();
        });

    [ToolMeta("mail_folder_create", "mail", Access.Write)]
    [Description("Create a mail folder.")]
    public static Task<string> MailFolderCreate(
        [Description("New folder name")] string name,
        [Description("Parent folder path or id; default the Inbox")] string? parent = null,
        [Description("Account for default-folder names")] string? account = null)
        => Sta.Run<string>(() =>
        {
            dynamic p = Ol.ResolveFolder(parent, account, Ol.Inbox);
            dynamic created = p.Folders.Add(name);
            return $"created {(string)p.Name}/{name} (id {Ol.FolderId(created)})";
        });

    // ------------------------------------------------------------------ helpers

    private static dynamic Compose(string[] to, string[]? cc, string[]? bcc, string subject, string body, bool html,
        string[]? attachments, string? account, string? importance, string? deferUntil)
    {
        if (to.Length == 0) throw new ToolError("At least one 'to' address is required.");
        var acct = Ol.FindAccount(account);
        // Outlook sends a POP/IMAP message as the account whose store holds the draft, and ignores
        // (or rejects, RPC_E_SERVERFAULT) SendUsingAccount for those accounts. So with an explicit
        // account the item is created inside that account's own Drafts folder.
        dynamic mi = acct is null ? Ol.App.CreateItem(0) : acct.DeliveryStore.GetDefaultFolder(Ol.Drafts).Items.Add(0);
        if (acct is not null)
        {
            try { mi.SendUsingAccount = acct; } catch (System.Runtime.InteropServices.COMException) { /* expected for POP/IMAP */ }
            string want = acct.DeliveryStore.StoreID;
            string have = mi.Parent.StoreID;
            if (want != have)
            {
                try { mi.Delete(); } catch { }
                throw new ToolError($"Could not create the message in {(string)acct.SmtpAddress}'s mailbox, so it would send from another account. Nothing was sent.");
            }
        }
        mi.To = string.Join(";", to);
        if (cc is { Length: > 0 }) mi.CC = string.Join(";", cc);
        if (bcc is { Length: > 0 }) mi.BCC = string.Join(";", bcc);
        mi.Subject = subject;
        if (html) mi.HTMLBody = body; else mi.Body = body;
        if (importance is not null) mi.Importance = Importance(importance);
        if (deferUntil is not null) mi.DeferredDeliveryTime = Ol.ParseDate(deferUntil, "defer_until");
        AddAttachments(mi, attachments);
        return mi;
    }

    private static void AddAttachments(dynamic mi, string[]? paths)
    {
        if (paths is null) return;
        foreach (var p in paths)
        {
            if (!File.Exists(p)) throw new ToolError($"Attachment not found: {p}");
            mi.Attachments.Add(p);
        }
    }

    private static string FinishComposed(dynamic mi, bool send, string verb)
    {
        var summary = RecipientSummary(mi);
        var subj = Ol.S(() => mi.Subject);
        if (send)
        {
            Guard(mi);
            mi.Send();
            return $"{verb} sent → {summary} | {subj}";
        }
        mi.Save();
        return $"{verb} draft: {Ol.Id(mi)} → {summary} | {subj}";
    }

    private static readonly Regex BodyTag = BodyTagRegex();

    private static void PrependBody(dynamic mi, string text, bool html)
    {
        var block = html ? text : "<div>" + System.Net.WebUtility.HtmlEncode(text).Replace("\r\n", "\n").Replace("\n", "<br>") + "</div><br>";
        string existing = Ol.S(() => mi.HTMLBody);
        if (existing.Length == 0) { mi.Body = text + "\r\n\r\n" + Ol.S(() => mi.Body); return; }
        mi.HTMLBody = BodyTag.IsMatch(existing)
            ? BodyTag.Replace(existing, m => m.Value + block, 1)
            : block + existing;
    }

    /// <summary>Fail-closed gate run before every send: recipients must resolve and pass the allow-list.</summary>
    private static void Guard(dynamic mi)
    {
        if (Current.Settings.DisableSend) throw new ToolError("Sending is disabled by configuration (OUTLOOK_MCP_DISABLE_SEND).");
        mi.Recipients.ResolveAll();
        var unresolved = new List<string>();
        var blocked = new List<string>();
        var allow = Current.Settings.AllowedDomains;
        foreach (dynamic r in mi.Recipients)
        {
            if (!(bool)r.Resolved) { unresolved.Add((string)r.Name); continue; }
            if (allow.Count == 0) continue;
            string smtp = Ol.RecipientSmtp(r);
            var domain = smtp.Contains('@') ? smtp[(smtp.LastIndexOf('@') + 1)..] : "";
            if (!allow.Contains(domain)) blocked.Add(smtp);
        }
        if (unresolved.Count > 0) throw new ToolError($"Unresolved recipient(s): {string.Join(", ", unresolved)}. Nothing was sent.");
        if (blocked.Count > 0) throw new ToolError($"Recipient domain not in OUTLOOK_MCP_ALLOWED_DOMAINS: {string.Join(", ", blocked)}. Nothing was sent.");
    }

    private static string RecipientSummary(dynamic mi)
    {
        var list = new List<string>();
        try
        {
            foreach (dynamic r in mi.Recipients)
            {
                var kind = Ol.Safe(() => (int)r.Type, 1);
                string addr = Ol.Safe(() => (bool)r.Resolved, false) ? Ol.RecipientSmtp(r) : (string)r.Name;
                list.Add(kind == 1 ? addr : kind == 2 ? $"cc:{addr}" : $"bcc:{addr}");
            }
        }
        catch { list.Add(Ol.S(() => mi.To)); }
        return string.Join(", ", list);
    }

    private static void RequireSent(dynamic mi)
    {
        if (Ol.Safe(() => (bool)mi.Sent == false, false))
            throw new ToolError("Drafts cannot be flagged for follow-up (Outlook only flags sent or received mail).");
    }

    private static int Importance(string value) => value.ToLowerInvariant() switch
    {
        "low" => 0, "normal" => 1, "high" => 2,
        _ => throw new ToolError("importance must be low, normal or high."),
    };

    private static string SenderSmtp(dynamic mi)
    {
        try
        {
            string type = Ol.S(() => mi.SenderEmailType);
            if (type == "EX")
            {
                var smtp = Ol.S(() => mi.Sender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E"));
                if (smtp.Length > 0) return smtp;
                var eu = mi.Sender?.GetExchangeUser();
                if (eu is not null) return eu.PrimarySmtpAddress;
            }
            return Ol.S(() => mi.SenderEmailAddress);
        }
        catch { return Ol.S(() => mi.SenderEmailAddress); }
    }

    private static void AppendRecipients(StringBuilder sb, dynamic mi, bool full)
    {
        if (!full)
        {
            sb.AppendLine($"to: {Fmt.OneLine(Ol.S(() => mi.To), 200)}");
            var cc = Ol.S(() => mi.CC);
            if (cc.Length > 0) sb.AppendLine($"cc: {Fmt.OneLine(cc, 200)}");
            return;
        }
        var groups = new Dictionary<int, List<string>> { [1] = [], [2] = [], [3] = [] };
        foreach (dynamic r in mi.Recipients)
        {
            int t = Ol.Safe(() => (int)r.Type, 1);
            groups[t].Add($"{(string)r.Name} <{Ol.RecipientSmtp(r)}>");
        }
        if (groups[1].Count > 0) sb.AppendLine($"to: {string.Join("; ", groups[1])}");
        if (groups[2].Count > 0) sb.AppendLine($"cc: {string.Join("; ", groups[2])}");
        if (groups[3].Count > 0) sb.AppendLine($"bcc: {string.Join("; ", groups[3])}");
    }

    private static void AppendAttachments(StringBuilder sb, dynamic mi)
    {
        int count = Ol.Safe(() => (int)mi.Attachments.Count, 0);
        if (count == 0) return;
        var parts = new List<string>();
        for (var i = 1; i <= count; i++)
        {
            dynamic a = mi.Attachments[i];
            var inline = Ol.S(() => a.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")).Length > 0;
            if (!inline) parts.Add($"{Ol.S(() => a.FileName)} ({Fmt.Size(Ol.Safe(() => (long)a.Size, 0L))})");
            Ol.Release(a);
        }
        if (parts.Count > 0) sb.AppendLine($"attachments: {string.Join("; ", parts)}");
    }

    [GeneratedRegex(@"<body[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex BodyTagRegex();
}
