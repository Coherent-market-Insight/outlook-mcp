using System.Collections.Concurrent;

namespace OutlookMcp.Com;

/// <summary>
/// Outlook's object model is single-threaded-apartment COM. Every COM call in this
/// process runs on this one dedicated STA thread, so COM objects never cross apartments
/// and tool calls are serialised without locks.
///
/// A COM call cannot be aborted. If Outlook blocks (a modal dialog, a sign-in prompt) the
/// caller is released after the configured timeout with an actionable message, work that
/// had not started yet is dropped (so a stale "send" can never fire later), and new calls
/// fail fast while the thread is still stuck instead of queueing behind it.
/// </summary>
internal static class Sta
{
    private sealed class Job(Action run) { public Action Run { get; } = run; public volatile bool Abandoned; }

    private static readonly BlockingCollection<Job> Queue = new();
    private static readonly Thread Worker = Start();
    private static int _jobs;
    private static long _busySince; // Environment.TickCount64 while a job runs, else 0

    private static Thread Start()
    {
        var t = new Thread(() =>
        {
            foreach (var job in Queue.GetConsumingEnumerable())
            {
                if (job.Abandoned) { Debug.Log("sta: dropped abandoned job"); continue; }
                Interlocked.Exchange(ref _busySince, Environment.TickCount64);
                try { job.Run(); }
                finally
                {
                    Interlocked.Exchange(ref _busySince, 0);
                    // COM wrappers are only released by the GC; collecting after each job keeps a long session flat.
                    if (++_jobs % 5 == 0) { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); }
                }
            }
        })
        { IsBackground = true, Name = "outlook-sta" };
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        return t;
    }

    public static async Task<T> Run<T>(Func<T> func)
    {
        _ = Worker;
        var limit = TimeSpan.FromSeconds(Current.Settings.TimeoutSeconds);

        var since = Interlocked.Read(ref _busySince);
        if (since != 0 && TimeSpan.FromMilliseconds(Environment.TickCount64 - since) > limit)
            throw new ToolError(Blocked(limit));

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var job = new Job(() =>
        {
            try { tcs.SetResult(Ol.WithRetry(func)); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        Queue.Add(job);

        var finished = await Task.WhenAny(tcs.Task, Task.Delay(limit));
        if (finished != tcs.Task)
        {
            job.Abandoned = true;
            throw new ToolError(Blocked(limit));
        }
        return await tcs.Task;
    }

    private static string Blocked(TimeSpan limit) =>
        $"Outlook did not answer within {limit.TotalSeconds:0}s. A dialog box is probably open in Outlook " +
        "(security prompt, send/receive, sign-in, 'work offline'). Close it and retry; nothing queued behind it will run.";
}
