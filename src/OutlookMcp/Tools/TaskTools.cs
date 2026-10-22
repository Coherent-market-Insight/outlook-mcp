using System.ComponentModel;
using System.Text;
using OutlookMcp.Com;

namespace OutlookMcp.Tools;

internal static class TaskTools
{
    private static readonly string[] StatusNames = ["not_started", "in_progress", "completed", "waiting", "deferred"];
    private static readonly DateTime NoDate = new(4501, 1, 1);

    [ToolMeta("tasks_list", "tasks", Access.Read)]
    [Description("List tasks, soonest due first. One line each: id, due date, subject, status, flags !=high.")]
    public static Task<string> TasksList(
        [Description("open (default), completed or all")] string status = "open",
        [Description("Subject or notes contain")] string? query = null,
        [Description("Only tasks due on or before this date")] string? due_before = null,
        [Description("Account whose tasks to list; default account")] string? account = null,
        [Description("Max rows (1-100)")] int limit = 25)
        => Sta.Run<string>(() =>
        {
            limit = Math.Clamp(limit, 1, 100);
            var open = status.ToLowerInvariant() switch
            {
                "open" => (bool?)true, "completed" or "done" => false, "all" => null,
                _ => throw new ToolError("status must be open, completed or all."),
            };
            var dueLimit = due_before is null ? (DateTime?)null : Ol.ParseDate(due_before, "due_before").Date.AddDays(1);
            dynamic items = Ol.DefaultFolder(Ol.Tasks, account).Items;
            items.Sort("[DueDate]");
            var sb = new StringBuilder();
            var shown = 0; var more = false;
            dynamic? it = items.GetFirst();
            while (it is not null)
            {
                var done = Ol.Safe(() => (bool)it.Complete, false);
                DateTime due = Ol.Safe(() => (DateTime)it.DueDate, NoDate);
                var subject = Ol.S(() => it.Subject);
                var matches = (open is null || open == !done)
                              && (dueLimit is null || (due != NoDate && due < dueLimit))
                              && (query is null || subject.Contains(query, StringComparison.OrdinalIgnoreCase) || Ol.S(() => it.Body).Contains(query, StringComparison.OrdinalIgnoreCase));
                if (matches)
                {
                    if (shown >= limit) { more = true; Ol.Release(it); break; }
                    sb.AppendLine(TaskLine(it)); shown++;
                }
                Ol.Release(it);
                it = items.GetNext();
            }
            var head = $"tasks: {shown} shown{(more ? ", more: raise limit or narrow filters" : "")}";
            return shown == 0 ? head : head + "\n" + sb.ToString().TrimEnd();
        });

    private static string TaskLine(dynamic t)
    {
        DateTime due = Ol.Safe(() => (DateTime)t.DueDate, NoDate);
        var status = StatusNames[Math.Clamp(Ol.Safe(() => (int)t.Status, 0), 0, 4)];
        return $"{Ol.Id(t)} {(due == NoDate ? "no due date" : "due " + Fmt.Day(due))} {Fmt.OneLine(Ol.S(() => t.Subject), 90)} ({status})" +
               Fmt.Flags(high: Ol.Safe(() => (int)t.Importance == 2, false));
    }

    [ToolMeta("task_save", "tasks", Access.Write)]
    [Description("Create a task, or update one when id is given (only the fields you pass change). Set status=completed to finish a task.")]
    public static Task<string> TaskSave(
        [Description("Existing task id to update; omit to create")] string? id = null,
        [Description("Title")] string? subject = null,
        [Description("Due date, ISO; empty string clears it")] string? due = null,
        [Description("Start date, ISO")] string? start = null,
        [Description("Notes")] string? body = null,
        [Description("not_started, in_progress, completed, waiting, deferred")] string? status = null,
        [Description("Percent complete 0-100")] int? percent = null,
        [Description("low, normal or high")] string? importance = null,
        [Description("Category names")] string[]? categories = null,
        [Description("Reminder datetime, ISO; empty string clears it")] string? reminder = null,
        [Description("Account to create it in; default account")] string? account = null)
        => Sta.Run<string>(() =>
        {
            dynamic t;
            if (id is null)
            {
                if (string.IsNullOrWhiteSpace(subject)) throw new ToolError("A new task needs a subject.");
                t = Ol.DefaultFolder(Ol.Tasks, account).Items.Add(3);
            }
            else t = Ol.GetItem(id);
            if (subject is not null) t.Subject = subject;
            if (due is not null) t.DueDate = due.Length == 0 ? NoDate : Ol.ParseDate(due, "due").Date;
            if (start is not null) t.StartDate = start.Length == 0 ? NoDate : Ol.ParseDate(start, "start").Date;
            if (body is not null) t.Body = body;
            if (status is not null)
            {
                var i = Array.IndexOf(StatusNames, status.ToLowerInvariant());
                if (i < 0) throw new ToolError("status must be not_started, in_progress, completed, waiting or deferred.");
                t.Status = i;
                if (i == 2) t.PercentComplete = 100;
            }
            if (percent is not null) t.PercentComplete = Math.Clamp(percent.Value, 0, 100);
            if (importance is not null)
                t.Importance = importance.ToLowerInvariant() switch { "low" => 0, "normal" => 1, "high" => 2, _ => throw new ToolError("importance must be low, normal or high.") };
            if (categories is not null) t.Categories = string.Join(", ", categories);
            if (reminder is not null)
            {
                if (reminder.Length == 0) t.ReminderSet = false;
                else { t.ReminderSet = true; t.ReminderTime = Ol.ParseDate(reminder, "reminder"); }
            }
            t.Save();
            return $"{(id is null ? "created" : "updated")}: {TaskLine(t)}";
        });

    [ToolMeta("task_delete", "tasks", Access.Delete)]
    [Description("Delete tasks (moved to Deleted Items).")]
    public static Task<string> TaskDelete([Description("Task ids")] string[] ids)
        => Sta.Run<string>(() =>
        {
            foreach (var id in ids) Ol.GetItem(id).Delete();
            return $"deleted {ids.Length}";
        });
}
