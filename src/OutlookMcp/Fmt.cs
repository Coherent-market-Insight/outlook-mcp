using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookMcp;

/// <summary>Compact, token-frugal text formatting for tool results.</summary>
internal static partial class Fmt
{
    public static string Date(DateTime d) => d.Year == DateTime.Now.Year
        ? d.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture)
        : d.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

    public static string Iso(DateTime d) => d.ToString("yyyy-MM-dd'T'HH:mm", CultureInfo.InvariantCulture);

    public static string Day(DateTime d) => d.ToString("yyyy-MM-dd ddd", CultureInfo.InvariantCulture);

    public static string Size(long bytes) => bytes switch
    {
        < 1024 => $"{bytes}B",
        < 1024 * 1024 => $"{bytes / 1024}KB",
        _ => $"{bytes / 1048576.0:0.#}MB",
    };

    /// <summary>One-line summary text: collapse whitespace, cap length.</summary>
    public static string OneLine(string? s, int max = 120)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var t = Whitespace().Replace(s, " ").Trim();
        return t.Length <= max ? t : t[..(max - 1)] + "…";
    }

    /// <summary>
    /// Cleans a message body for model consumption: removes link and mailto noise,
    /// collapses blank runs, optionally cuts the quoted reply chain.
    /// </summary>
    public static string CleanBody(string? body, bool keepLinks, bool keepQuoted)
    {
        if (string.IsNullOrEmpty(body)) return "";
        var t = body.Replace("\r\n", "\n").Replace('\r', '\n');
        if (!keepLinks) t = AngleLink().Replace(t, "");
        if (!keepQuoted) t = CutQuoted(t);
        t = Banner().Replace(t, "");
        t = TrailingSpace().Replace(t, "\n");
        t = Disclaimer(t);
        t = BlankRuns().Replace(t, "\n\n");
        t = Undouble(t);
        return t.Trim();
    }

    /// <summary>Cuts a trailing legal footer ("CONFIDENTIAL...", "DISCLAIMER:") when it sits in the last part of the text.</summary>
    private static string Disclaimer(string t)
    {
        var m = DisclaimerStart().Match(t);
        return m.Success && m.Index > t.Length * 0.4 ? t[..m.Index].TrimEnd() : t;
    }

    /// <summary>Mail clients often double-space plain text; if nearly every line is followed by a blank, drop the blanks.</summary>
    private static string Undouble(string t)
    {
        var lines = t.Split('\n');
        var filled = 0; var gaps = 0;
        for (var i = 0; i < lines.Length - 1; i++)
            if (lines[i].Trim().Length > 0) { filled++; if (lines[i + 1].Trim().Length == 0) gaps++; }
        return filled >= 6 && gaps >= filled * 0.8
            ? string.Join('\n', lines.Where(l => l.Trim().Length > 0))
            : t;
    }

    public static string CutQuoted(string t)
    {
        var m = QuoteStart().Match(t);
        return m.Success && m.Index > 0 ? t[..m.Index].TrimEnd() : t;
    }

    public static string StripHtml(string html)
    {
        var t = HtmlBlocks().Replace(html, " ");
        t = HtmlBreaks().Replace(t, "\n");
        t = HtmlTags().Replace(t, "");
        return System.Net.WebUtility.HtmlDecode(t);
    }

    /// <summary>Returns a window of text with a precise continuation hint.</summary>
    public static string Window(string text, int offset, int max, string offsetParam)
    {
        if (offset >= text.Length) return offset == 0 ? "" : "(end of body)";
        var end = Math.Min(text.Length, offset + max);
        var slice = text[offset..end];
        return end < text.Length
            ? $"{slice}\n…[{text.Length - end} more chars; {offsetParam}={end}]"
            : slice;
    }

    public static string Flags(bool unread = false, bool high = false, bool attach = false, bool flagged = false, bool recurring = false)
    {
        var sb = new StringBuilder();
        if (unread) sb.Append('U');
        if (high) sb.Append('!');
        if (attach) sb.Append('@');
        if (flagged) sb.Append('F');
        if (recurring) sb.Append('R');
        return sb.Length == 0 ? "" : $" [{sb}]";
    }

    [GeneratedRegex(@"\s+")] private static partial Regex Whitespace();
    [GeneratedRegex(@"[ \t]*<(?:https?://|mailto:|ftp://)[^>\s]*>")] private static partial Regex AngleLink();
    [GeneratedRegex(@"[ \t]+\n")] private static partial Regex TrailingSpace();
    [GeneratedRegex(@"\A\s*SECURITY WARNING!.*?(?:\n\s*\n|\z)", RegexOptions.Singleline)] private static partial Regex Banner();
    [GeneratedRegex(@"(?im)^\s*(?:confidential(?:ity)?(?: and privileged| notice)?\b|disclaimer\s*:|this e-?mail and any attachments are confidential)")] private static partial Regex DisclaimerStart();
    [GeneratedRegex(@"\n{3,}")] private static partial Regex BlankRuns();
    [GeneratedRegex(@"(?im)^(?:-{2,}\s*original message\s*-{2,}|_{10,}|from:.*\n(?:.*\n){0,2}?\s*(?:sent|date):|on .{5,120} wrote:\s*$)")]
    private static partial Regex QuoteStart();
    [GeneratedRegex(@"(?is)<(?:style|script|head)\b.*?</(?:style|script|head)>")] private static partial Regex HtmlBlocks();
    [GeneratedRegex(@"(?i)<\s*(?:br\s*/?|/p|/div|/tr|/li|/h\d)\s*>")] private static partial Regex HtmlBreaks();
    [GeneratedRegex(@"<[^>]+>")] private static partial Regex HtmlTags();
}
