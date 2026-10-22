using System.ComponentModel;
using System.Text;
using OutlookMcp.Com;

namespace OutlookMcp.Tools;

internal static class CalendarTools
{
    private static readonly string[] ShowAsNames = ["free", "tentative", "busy", "oof", "elsewhere"];
    private static readonly string[] ResponseNames = ["none", "organizer", "tentative", "accepted", "declined", "no response"];

    // ------------------------------------------------------------------ calendars

    [ToolMeta("calendar_list", "calendar", Access.Read)]
    [Description("List calendars across all accounts. Use a returned name as the `calendar` argument of other calendar tools.")]
    public static Task<string> CalendarList() => Sta.Run<string>(() =>
    {
        var sb = new StringBuilder();
        foreach (dynamic st in Ol.Ns.Stores)
        {
            try
            {
                dynamic root = st.GetRootFolder();
                string storeName = st.DisplayName;
                dynamic def = st.GetDefaultFolder(Ol.Calendar);
                sb.AppendLine($"{storeName}{(storeName == Ol.S(() => Ol.Ns.DefaultStore.DisplayName) ? " (default)" : "")}: {(string)def.Name}, {(int)def.Items.Count} items");
                foreach (dynamic f in root.Folders)
                    if ((int)f.DefaultItemType == 1 && (string)f.EntryID != (string)def.EntryID)
                        sb.AppendLine($"  {storeName}/{(string)f.Name}: {(int)f.Items.Count} items");
            }
            catch { /* stores without a calendar (public folders, archives) */ }
        }
        return sb.ToString().TrimEnd();
    });

    internal static dynamic ResolveCalendar(string? spec)
    {
        if (string.IsNullOrWhiteSpace(spec) || spec.Equals("default", StringComparison.OrdinalIgnoreCase))
            return Ol.Ns.GetDefaultFolder(Ol.Calendar);
        var s = System.Text.RegularExpressions.Regex.Replace(spec.Trim(), @" - Calendar$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        foreach (dynamic st in Ol.Ns.Stores)
        {
            try
            {
                string storeName = st.DisplayName;
                dynamic def = st.GetDefaultFolder(Ol.Calendar);
                if (storeName.Equals(s, StringComparison.OrdinalIgnoreCase) || storeName.Equals(spec, StringComparison.OrdinalIgnoreCase)) return def;
                foreach (dynamic f in st.GetRootFolder().Folders)
                    if ((int)f.DefaultItemType == 1 &&
                        ((string)f.Name).Equals(s, StringComparison.OrdinalIgnoreCase) ||
                        $"{storeName}/{(string)f.Name}".Equals(s, StringComparison.OrdinalIgnoreCase)) return f;
            }
            catch { }
        }
        dynamic? acct = null;
        try { acct = Ol.FindAccount(spec); } catch (ToolError) { }
        return acct is not null ? acct.DeliveryStore.GetDefaultFolder(Ol.Calendar) : throw new ToolError($"Calendar '{spec}' not found. Call calendar_list.");
    }

    // ------------------------------------------------------------------ list / get

    [ToolMeta("calendar_events", "calendar", Access.Read)]
    [Description("Events in a date range, in time order, recurring meetings expanded. One line each: id, time, subject, location, flags R=recurring M=meeting. For one occurrence of a recurring event pass its start as occurrence_start to other tools.")]
    public static Task<string> CalendarEvents(
        [Description("Range start, ISO date or datetime")] string start,
        [Description("Range end; default 7 days after start")] string? end = null,
        [Description("Calendar name from calendar_list; default calendar otherwise")] string? calendar = null,
        [Description("Subject or location contains")] string? query = null,
        [Description("Max events (1-200)")] int limit = 50)
        => Sta.Run<string>(() =>
        {
            limit = Math.Clamp(limit, 1, 200);
            var from = Ol.ParseDate(start, "start");
            var to = end is null ? from.Date.AddDays(7) : Ol.ParseDate(end, "end");
            if (end is not null && to.TimeOfDay == TimeSpan.Zero && to.Date == to) to = to.AddDays(1);
            if (to <= from) throw new ToolError("'end' must be after 'start'.");

            dynamic folder = ResolveCalendar(calendar);
            dynamic items = folder.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            dynamic hits = items.Restrict($"[Start] < '{Ol.JetDate(to)}' AND [End] > '{Ol.JetDate(from)}'");

            var sb = new StringBuilder();
            var shown = 0; var more = false;
            dynamic? it = hits.GetFirst();
            while (it is not null)
            {
                var subject = Ol.S(() => it.Subject);
                var loc = Ol.S(() => it.Location);
                if (query is null || subject.Contains(query, StringComparison.OrdinalIgnoreCase) || loc.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    if (shown >= limit) { more = true; Ol.Release(it); break; }
                    sb.AppendLine(EventLine(it, folder));
                    shown++;
                }
                Ol.Release(it);
                it = hits.GetNext();
            }
            var head = $"{(string)folder.Name}: {shown} event(s) {Fmt.Day(from)} → {Fmt.Day(to.AddMinutes(-1))}{(more ? ", more: narrow the range or raise limit" : "")}";
            return shown == 0 ? head : head + "\n" + sb.ToString().TrimEnd();
        });

    private static string EventLine(dynamic ap, dynamic? folder = null, string? knownId = null)
    {
        DateTime s = Ol.Safe(() => (DateTime)ap.Start, DateTime.MinValue);
        DateTime e = Ol.Safe(() => (DateTime)ap.End, s);
        var allDay = Ol.Safe(() => (bool)ap.AllDayEvent, false);
        var when = allDay
            ? $"{Fmt.Day(s)} all-day"
            : s.Date == e.Date ? $"{Fmt.Day(s)} {s:HH:mm}-{e:HH:mm}" : $"{Fmt.Day(s)} {s:HH:mm} → {Fmt.Day(e)} {e:HH:mm}";
        var loc = Ol.S(() => ap.Location);
        var busy = Ol.Safe(() => (int)ap.BusyStatus, 2);
        var meeting = Ol.Safe(() => (int)ap.MeetingStatus, 0) != 0;
        var recurring = Ol.Safe(() => (bool)ap.IsRecurring, false) || Ol.Safe(() => (int)ap.RecurrenceState, 0) != 0;
        return $"{knownId ?? Ol.Id(ap, folder)} {when} {Fmt.OneLine(Ol.S(() => ap.Subject), 90)}" +
               (loc.Length > 0 ? $" @{Fmt.OneLine(loc, 40)}" : "") +
               (busy != 2 ? $" ({ShowAsNames[Math.Clamp(busy, 0, 4)]})" : "") +
               Fmt.Flags(recurring: recurring) + (meeting ? " [M]" : "");
    }

    [ToolMeta("calendar_event_get", "calendar", Access.Read)]
    [Description("Full details of one event: time, location, organizer, attendees with their responses, recurrence, reminder, body.")]
    public static Task<string> CalendarEventGet(
        [Description("Event id")] string id,
        [Description("Start of one occurrence when the event recurs")] string? occurrence_start = null,
        [Description("Max body characters")] int max_chars = 2000)
        => Sta.Run<string>(() =>
        {
            dynamic ap = OpenEvent(id, occurrence_start);
            var sb = new StringBuilder();
            sb.AppendLine(EventLine(ap, null, "").Trim());
            sb.AppendLine($"id: {id}");
            var org = Ol.S(() => ap.Organizer);
            if (org.Length > 0) sb.AppendLine($"organizer: {org}");
            var reminder = Ol.Safe(() => (bool)ap.ReminderSet, false) ? $"{(int)ap.ReminderMinutesBeforeStart} min before" : "none";
            sb.AppendLine($"reminder: {reminder}");
            var cats = Ol.S(() => ap.Categories);
            if (cats.Length > 0) sb.AppendLine($"categories: {cats}");
            var rec = Ol.Safe(() => (int)ap.RecurrenceState, 0);
            if (rec != 0) sb.AppendLine($"recurrence: {DescribeRecurrence(ap)}");
            var attendees = new List<string>();
            foreach (dynamic r in ap.Recipients)
            {
                var kind = Ol.Safe(() => (int)r.Type, 1) switch { 2 => " (optional)", 3 => " (resource)", _ => "" };
                var resp = ResponseNames[Math.Clamp(Ol.Safe(() => (int)r.MeetingResponseStatus, 0), 0, 5)];
                attendees.Add($"{(string)r.Name} <{Ol.RecipientSmtp(r)}>{kind}: {resp}");
            }
            if (attendees.Count > 0) sb.AppendLine("attendees:\n  " + string.Join("\n  ", attendees));
            var body = Fmt.CleanBody(Ol.S(() => ap.Body), false, true);
            if (body.Length > 0) sb.Append("--\n" + Fmt.Window(body, 0, max_chars, "max_chars"));
            return sb.ToString().TrimEnd();
        });

    // ------------------------------------------------------------------ create / update / delete

    [ToolMeta("calendar_event_create", "calendar", Access.Write)]
    [Description("Create an event or meeting. With attendees it is saved unsent unless send_invites=true. Supports all-day events, reminders, show-as and simple recurrence.")]
    public static Task<string> CalendarEventCreate(
        [Description("Title")] string subject,
        [Description("Start, ISO datetime (all-day: ISO date)")] string start,
        [Description("End, ISO datetime; default start + duration_minutes. All-day: last day, inclusive")] string? end = null,
        [Description("Minutes when end is omitted; default 30")] int duration_minutes = 30,
        [Description("true: all-day event")] bool all_day = false,
        [Description("Location")] string? location = null,
        [Description("Description")] string? body = null,
        [Description("Required attendee addresses")] string[]? attendees = null,
        [Description("Optional attendee addresses")] string[]? optional_attendees = null,
        [Description("true: send meeting invitations now")] bool send_invites = false,
        [Description("free, tentative, busy (default), oof, elsewhere")] string? show_as = null,
        [Description("Reminder minutes before; 0 for at start, -1 for none; default per Outlook")] int? reminder_minutes = null,
        [Description("Category names")] string[]? categories = null,
        [Description("Repeat: daily, weekly, monthly or yearly")] string? repeat = null,
        [Description("Repeat every N units; default 1")] int repeat_interval = 1,
        [Description("Weekly repeat days, e.g. [\"mon\",\"wed\"]")] string[]? repeat_days = null,
        [Description("Repeat until this date")] string? repeat_until = null,
        [Description("Repeat this many times")] int? repeat_count = null,
        [Description("Calendar name from calendar_list")] string? calendar = null)
        => Sta.Run<string>(() =>
        {
            var s = Ol.ParseDate(start, "start");
            DateTime e;
            if (all_day) { s = s.Date; e = (end is null ? s : Ol.ParseDate(end, "end").Date).AddDays(1); }
            else e = end is null ? s.AddMinutes(Math.Max(1, duration_minutes)) : Ol.ParseDate(end, "end");
            if (e <= s) throw new ToolError("'end' must be after 'start'.");

            dynamic folder = ResolveCalendar(calendar);
            dynamic ap = folder.Items.Add(1);
            ap.Subject = subject;
            ap.AllDayEvent = all_day;
            ap.Start = s;
            ap.End = e;
            if (location is not null) ap.Location = location;
            if (body is not null) ap.Body = body;
            if (show_as is not null) ap.BusyStatus = ShowAs(show_as);
            if (reminder_minutes is not null)
            {
                if (reminder_minutes < 0) ap.ReminderSet = false;
                else { ap.ReminderSet = true; ap.ReminderMinutesBeforeStart = reminder_minutes; }
            }
            if (categories is { Length: > 0 }) ap.Categories = string.Join(", ", categories);
            if (repeat is not null) SetRecurrence(ap, repeat, repeat_interval, repeat_days, repeat_until, repeat_count, s, e);

            var invitees = (attendees?.Length ?? 0) + (optional_attendees?.Length ?? 0);
            if (invitees > 0)
            {
                ap.MeetingStatus = 1;
                foreach (var a in attendees ?? []) { dynamic r = ap.Recipients.Add(a); r.Type = 1; }
                foreach (var a in optional_attendees ?? []) { dynamic r = ap.Recipients.Add(a); r.Type = 2; }
                ap.Recipients.ResolveAll();
                foreach (dynamic r in ap.Recipients)
                    if (!(bool)r.Resolved) throw new ToolError($"Unresolved attendee: {(string)r.Name}. Nothing was created.");
            }
            if (send_invites && invitees > 0) ap.Send(); else ap.Save();
            return $"{(send_invites && invitees > 0 ? "invites sent" : invitees > 0 ? "meeting saved (invites NOT sent)" : "event created")}: {EventLine(ap)}";
        });

    [ToolMeta("calendar_event_update", "calendar", Access.Write)]
    [Description("Change an event. Only the fields you pass change. For a recurring event pass occurrence_start to change a single occurrence.")]
    public static Task<string> CalendarEventUpdate(
        [Description("Event id")] string id,
        [Description("Start of one occurrence when the event recurs")] string? occurrence_start = null,
        [Description("New title")] string? subject = null,
        [Description("New start, ISO; the duration is kept unless end is also given")] string? start = null,
        [Description("New end, ISO")] string? end = null,
        [Description("New location")] string? location = null,
        [Description("New description")] string? body = null,
        [Description("free, tentative, busy, oof, elsewhere")] string? show_as = null,
        [Description("Reminder minutes before; -1 for none")] int? reminder_minutes = null,
        [Description("Replace categories")] string[]? categories = null,
        [Description("true: send the update to attendees (meetings you organise)")] bool send_update = false)
        => Sta.Run<string>(() =>
        {
            dynamic ap = OpenEvent(id, occurrence_start);
            if (subject is not null) ap.Subject = subject;
            if (start is not null || end is not null)
            {
                DateTime oldS = ap.Start, oldE = ap.End;
                var ns = start is null ? oldS : Ol.ParseDate(start, "start");
                var ne = end is not null ? Ol.ParseDate(end, "end") : ns + (oldE - oldS);
                if (ne <= ns) throw new ToolError("'end' must be after 'start'.");
                ap.Start = ns; // Outlook moves End with Start to keep the old duration...
                ap.End = ne;   // ...so End is always set explicitly afterwards
            }
            if (location is not null) ap.Location = location;
            if (body is not null) ap.Body = body;
            if (show_as is not null) ap.BusyStatus = ShowAs(show_as);
            if (reminder_minutes is not null)
            {
                if (reminder_minutes < 0) ap.ReminderSet = false;
                else { ap.ReminderSet = true; ap.ReminderMinutesBeforeStart = reminder_minutes; }
            }
            if (categories is not null) ap.Categories = string.Join(", ", categories);
            var isMeeting = Ol.Safe(() => (int)ap.MeetingStatus, 0) == 1 && Ol.Safe(() => (int)ap.Recipients.Count, 0) > 0;
            if (send_update && isMeeting) ap.Send(); else ap.Save();
            return $"updated: {EventLine(ap, null, id)}" + (isMeeting && !send_update ? "\n(attendees were not notified)" : "");
        });

    [ToolMeta("calendar_event_delete", "calendar", Access.Delete)]
    [Description("Delete an event, or one occurrence of a recurring event. With cancel_attendees=true a meeting you organise is cancelled and attendees are notified.")]
    public static Task<string> CalendarEventDelete(
        [Description("Event id")] string id,
        [Description("Start of one occurrence when the event recurs")] string? occurrence_start = null,
        [Description("true: send cancellations to attendees")] bool cancel_attendees = false)
        => Sta.Run<string>(() =>
        {
            dynamic ap = OpenEvent(id, occurrence_start);
            var subject = Ol.S(() => ap.Subject);
            var isMeeting = Ol.Safe(() => (int)ap.MeetingStatus, 0) == 1 && Ol.Safe(() => (int)ap.Recipients.Count, 0) > 0;
            if (cancel_attendees && isMeeting) { ap.MeetingStatus = 5; ap.Send(); return $"cancelled and notified: {subject}"; }
            ap.Delete();
            return $"deleted: {subject}" + (isMeeting ? "\n(attendees were not notified)" : "");
        });

    [ToolMeta("calendar_respond", "calendar", Access.Send)]
    [Description("Respond to a meeting invitation: accept, tentative or decline. Pass the invitation's message id from the inbox, or an event id.")]
    public static Task<string> CalendarRespond(
        [Description("Meeting request message id, or event id")] string id,
        [Description("accept, tentative or decline")] string response,
        [Description("true: send the response to the organiser (default)")] bool send_response = true,
        [Description("Optional note to the organiser")] string? comment = null)
        => Sta.Run<string>(() =>
        {
            var code = response.ToLowerInvariant() switch
            {
                "accept" or "accepted" => 3, "tentative" => 2, "decline" or "declined" => 4,
                _ => throw new ToolError("response must be accept, tentative or decline."),
            };
            dynamic item = Ol.GetItem(id);
            dynamic ap = Ol.Safe(() => (int)item.Class, 0) == 53 ? item.GetAssociatedAppointment(true) : item;
            string subject = Ol.S(() => ap.Subject);
            dynamic reply = ap.Respond(code, true, false);
            if (reply is not null && send_response)
            {
                if (!string.IsNullOrEmpty(comment)) reply.Body = comment;
                reply.Send();
            }
            return $"{response.ToLowerInvariant()}{(send_response ? "" : " (response not sent)")}: {subject}";
        });

    // ------------------------------------------------------------------ availability

    [ToolMeta("calendar_free_busy", "calendar", Access.Read)]
    [Description("Find free slots in working hours. Uses your calendar, plus other people's free/busy when attendees are given (Exchange/M365 only; others report no data).")]
    public static Task<string> CalendarFreeBusy(
        [Description("Range start, ISO date or datetime")] string start,
        [Description("Range end; default 7 days after start")] string? end = null,
        [Description("Slot length in minutes; default 30")] int duration_minutes = 30,
        [Description("Other people's addresses to include")] string[]? attendees = null,
        [Description("Work day start hour 0-23; default 9")] int work_start = 9,
        [Description("Work day end hour 1-24; default 17")] int work_end = 17,
        [Description("true: include Saturdays and Sundays")] bool weekends = false,
        [Description("Your calendar name; default calendar")] string? calendar = null,
        [Description("Max windows to list")] int limit = 20)
        => Sta.Run<string>(() =>
        {
            var from = Ol.ParseDate(start, "start");
            var to = end is null ? from.Date.AddDays(7) : Ol.ParseDate(end, "end");
            if (end is not null && to.TimeOfDay == TimeSpan.Zero) to = to.AddDays(1);
            if (to <= from) throw new ToolError("'end' must be after 'start'.");
            if (work_end <= work_start) throw new ToolError("work_end must be after work_start.");
            duration_minutes = Math.Clamp(duration_minutes, 5, 720);

            const int slot = 30;
            var origin = from.Date;
            var totalSlots = (int)Math.Ceiling((to - origin).TotalMinutes / slot);
            if (totalSlots > 48 * 62) throw new ToolError("Range too long; use at most 60 days.");
            var busy = new bool[totalSlots];

            // own calendar
            dynamic items = ResolveCalendar(calendar).Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            dynamic hits = items.Restrict($"[Start] < '{Ol.JetDate(to)}' AND [End] > '{Ol.JetDate(origin)}'");
            dynamic? it = hits.GetFirst();
            while (it is not null)
            {
                var status = Ol.Safe(() => (int)it.BusyStatus, 2);
                if (status is 1 or 2 or 3)
                {
                    DateTime s = it.Start, e = it.End;
                    MarkBusy(busy, origin, s, e, slot);
                }
                Ol.Release(it);
                it = hits.GetNext();
            }

            // other people
            var notes = new List<string>();
            foreach (var who in attendees ?? [])
            {
                dynamic rec = Ol.Ns.CreateRecipient(who);
                rec.Resolve();
                var fb = (bool)rec.Resolved ? Ol.S(() => rec.FreeBusy(origin, slot, true)) : "";
                if (fb.Length == 0) { notes.Add($"no free/busy data for {who}"); continue; }
                for (var i = 0; i < Math.Min(fb.Length, totalSlots); i++)
                    if (fb[i] is '1' or '2' or '3') busy[i] = true;
            }

            var windows = new List<string>();
            for (var day = from.Date; day < to.Date.AddDays(to.TimeOfDay == TimeSpan.Zero ? 0 : 1); day = day.AddDays(1))
            {
                if (!weekends && day.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
                DateTime? runStart = null;
                for (var t = day.AddHours(work_start); t <= day.AddHours(work_end); t = t.AddMinutes(slot))
                {
                    var inRange = t >= from && t < to && t < day.AddHours(work_end);
                    var idx = (int)((t - origin).TotalMinutes / slot);
                    var free = inRange && idx >= 0 && idx < totalSlots && !busy[idx];
                    if (free) runStart ??= t;
                    else if (runStart is not null)
                    {
                        if ((t - runStart.Value).TotalMinutes >= duration_minutes) windows.Add($"{Fmt.Day(day)} {runStart:HH:mm}-{t:HH:mm}");
                        runStart = null;
                    }
                }
            }
            var shown = windows.Take(Math.Clamp(limit, 1, 100)).ToList();
            var sb = new StringBuilder();
            sb.AppendLine($"free windows ≥{duration_minutes} min, {work_start:00}:00-{work_end:00}:00 ({windows.Count}{(windows.Count > shown.Count ? $", showing {shown.Count}" : "")}):");
            sb.AppendLine(shown.Count == 0 ? "  none" : string.Join("\n", shown.Select(w => "  " + w)));
            foreach (var n in notes) sb.AppendLine($"note: {n}");
            return sb.ToString().TrimEnd();
        });

    private static void MarkBusy(bool[] busy, DateTime origin, DateTime s, DateTime e, int slot)
    {
        var first = Math.Max(0, (int)Math.Floor((s - origin).TotalMinutes / slot));
        var last = Math.Min(busy.Length - 1, (int)Math.Ceiling((e - origin).TotalMinutes / slot) - 1);
        for (var i = first; i <= last; i++) busy[i] = true;
    }

    // ------------------------------------------------------------------ helpers

    private static dynamic OpenEvent(string id, string? occurrenceStart)
    {
        dynamic item = Ol.GetItem(id);
        if (occurrenceStart is null) return item;
        if (!Ol.Safe(() => (bool)item.IsRecurring, false)) return item;
        var when = Ol.ParseDate(occurrenceStart, "occurrence_start");
        try { return item.GetRecurrencePattern().GetOccurrence(when); }
        catch (System.Runtime.InteropServices.COMException)
        {
            throw new ToolError($"No occurrence starts at {Fmt.Iso(when)}. Copy the start time shown by calendar_events exactly.");
        }
    }

    private static int ShowAs(string value)
    {
        var i = Array.IndexOf(ShowAsNames, value.ToLowerInvariant() switch { "outofoffice" or "out of office" => "oof", "workingelsewhere" => "elsewhere", var v => v });
        return i >= 0 ? i : throw new ToolError("show_as must be free, tentative, busy, oof or elsewhere.");
    }

    private static void SetRecurrence(dynamic ap, string repeat, int interval, string[]? days, string? until, int? count, DateTime start, DateTime end)
    {
        dynamic p = ap.GetRecurrencePattern();
        interval = Math.Max(1, interval);
        switch (repeat.ToLowerInvariant())
        {
            case "daily": p.RecurrenceType = 0; p.Interval = interval; break;
            case "weekly":
                p.RecurrenceType = 1; p.Interval = interval;
                p.DayOfWeekMask = days is { Length: > 0 } ? DayMask(days) : DayMask([start.DayOfWeek.ToString()[..3]]);
                break;
            case "monthly": p.RecurrenceType = 2; p.Interval = interval; p.DayOfMonth = start.Day; break;
            case "yearly": p.RecurrenceType = 5; p.MonthOfYear = start.Month; p.DayOfMonth = start.Day; break;
            default: throw new ToolError("repeat must be daily, weekly, monthly or yearly.");
        }
        p.PatternStartDate = start.Date;
        p.StartTime = new DateTime(1899, 12, 30) + start.TimeOfDay;
        p.EndTime = new DateTime(1899, 12, 30) + end.TimeOfDay;
        if (until is not null) p.PatternEndDate = Ol.ParseDate(until, "repeat_until").Date;
        else if (count is not null) p.Occurrences = count;
        else p.NoEndDate = true;
    }

    private static int DayMask(string[] days)
    {
        var mask = 0;
        foreach (var d in days)
        {
            mask |= d.Trim().ToLowerInvariant()[..Math.Min(3, d.Trim().Length)] switch
            {
                "sun" => 1, "mon" => 2, "tue" => 4, "wed" => 8, "thu" => 16, "fri" => 32, "sat" => 64,
                _ => throw new ToolError($"Unknown weekday '{d}'. Use mon, tue, wed, thu, fri, sat, sun."),
            };
        }
        return mask;
    }

    private static string DescribeRecurrence(dynamic ap)
    {
        try
        {
            dynamic p = ap.GetRecurrencePattern();
            var kind = (int)p.RecurrenceType switch { 0 => "daily", 1 => "weekly", 2 or 3 => "monthly", 5 or 6 => "yearly", _ => "custom" };
            var every = (int)p.Interval > 1 ? $" every {(int)p.Interval}" : "";
            var tail = (bool)p.NoEndDate ? ", no end" : $", until {Fmt.Day((DateTime)p.PatternEndDate)}";
            return kind + every + tail;
        }
        catch { return "recurring"; }
    }
}
