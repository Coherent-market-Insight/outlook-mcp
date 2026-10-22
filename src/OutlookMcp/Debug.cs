using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OutlookMcp;

/// <summary>
/// Debug mode (--debug or OUTLOOK_MCP_DEBUG=1): one log line per tool call on stderr and,
/// optionally, in a file (--log-file / OUTLOOK_MCP_LOG_FILE). Message bodies, subjects and
/// addresses are never logged: string arguments are reduced to their length.
/// </summary>
internal static class Debug
{
    private static readonly object Gate = new();
    public static bool Enabled => Current.Settings.Debug;

    public static void Log(string line)
    {
        if (!Enabled) return;
        var text = $"{DateTime.Now:HH:mm:ss.fff} {line}";
        lock (Gate)
        {
            Console.Error.WriteLine(text);
            var file = Current.Settings.LogFile;
            if (file is not null)
                try { File.AppendAllText(file, text + Environment.NewLine, Encoding.UTF8); } catch { /* logging must never break a call */ }
        }
    }

    public static string Summarise(IDictionary<string, JsonElement>? args)
    {
        if (args is null || args.Count == 0) return "{}";
        var parts = args.Select(kv => $"{kv.Key}={Short(kv.Value)}");
        return "{" + string.Join(", ", parts) + "}";
    }

    private static string Short(JsonElement v) => v.ValueKind switch
    {
        JsonValueKind.String => IsPlainToken(v.GetString()!) ? v.GetString()! : $"str({v.GetString()!.Length})",
        JsonValueKind.Array => $"[{v.GetArrayLength()}]",
        JsonValueKind.Object => "{…}",
        JsonValueKind.Null => "null",
        _ => v.ToString(),
    };

    // short non-sensitive tokens (ids, folder names, flags) are shown; anything free-text is only measured
    private static bool IsPlainToken(string s) =>
        s.Length <= 24 && s.All(c => char.IsLetterOrDigit(c) || c is ':' or '~' or '_' or '-' or '/' or '.');

    public static string Elapsed(Stopwatch sw) => $"{sw.ElapsedMilliseconds} ms";
}
