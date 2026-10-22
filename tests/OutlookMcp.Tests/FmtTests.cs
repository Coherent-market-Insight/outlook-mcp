using OutlookMcp;

namespace OutlookMcp.Tests;

public class FmtTests
{
    [Fact]
    public void CleanBody_removes_gateway_banner_and_link_noise()
    {
        var raw = "SECURITY WARNING! This is an external email. Do not click links.\r\n\r\nHello Sam,\r\nPlease see <https://example.com/track?id=1> and mail <mailto:a@b.com>.";
        var clean = Fmt.CleanBody(raw, keepLinks: false, keepQuoted: true);
        Assert.StartsWith("Hello Sam,", clean);
        Assert.DoesNotContain("http", clean);
        Assert.DoesNotContain("mailto", clean);
    }

    [Fact]
    public void CleanBody_keeps_links_when_asked()
    {
        var clean = Fmt.CleanBody("See <https://example.com/x>", keepLinks: true, keepQuoted: true);
        Assert.Contains("https://example.com/x", clean);
    }

    [Theory]
    [InlineData("Thanks!\n\nFrom: Bob <b@x.com>\nSent: Monday\nTo: me\nSubject: hi\n\nold text")]
    [InlineData("Thanks!\n\n-----Original Message-----\nFrom: Bob\nold text")]
    [InlineData("Thanks!\n\nOn Mon, 3 Oct 2026 at 10:00, Bob <b@x.com> wrote:\n> old text")]
    public void CutQuoted_drops_the_reply_chain(string raw)
    {
        var clean = Fmt.CleanBody(raw, keepLinks: false, keepQuoted: false);
        Assert.Equal("Thanks!", clean);
    }

    [Fact]
    public void CleanBody_keeps_the_quoted_chain_when_asked()
    {
        var raw = "Thanks!\n\n-----Original Message-----\nFrom: Bob\nold text";
        Assert.Contains("old text", Fmt.CleanBody(raw, false, keepQuoted: true));
    }

    [Fact]
    public void CleanBody_undoubles_double_spaced_text()
    {
        var raw = string.Join("\n\n", Enumerable.Range(1, 10).Select(i => $"line {i}"));
        var clean = Fmt.CleanBody(raw, false, true);
        Assert.DoesNotContain("\n\n", clean);
        Assert.Equal(10, clean.Split('\n').Length);
    }

    [Fact]
    public void CleanBody_keeps_real_paragraphs()
    {
        var raw = "First paragraph line one\nline two\n\nSecond paragraph\nmore text";
        Assert.Contains("\n\n", Fmt.CleanBody(raw, false, true));
    }

    [Fact]
    public void CleanBody_cuts_a_trailing_legal_footer_but_not_one_at_the_top()
    {
        var body = "Hi Raj,\nHere is the report you asked for with all the details you need to review.\nRegards\nAmeya\n\nCONFIDENTIAL AND PRIVILEGED: This message is confidential.";
        var clean = Fmt.CleanBody(body, false, true);
        Assert.EndsWith("Ameya", clean);
        var top = "CONFIDENTIAL: draft contract attached\nPlease review the clauses below carefully before Friday.";
        Assert.StartsWith("CONFIDENTIAL", Fmt.CleanBody(top, false, true));
    }

    [Fact]
    public void CleanBody_handles_null_and_empty()
    {
        Assert.Equal("", Fmt.CleanBody(null, false, false));
        Assert.Equal("", Fmt.CleanBody("", false, false));
    }

    [Theory]
    [InlineData("a   b\n\t c", 50, "a b c")]
    [InlineData("abcdefghij", 5, "abcd…")]
    [InlineData(null, 5, "")]
    public void OneLine_collapses_whitespace_and_caps_length(string? input, int max, string expected)
        => Assert.Equal(expected, Fmt.OneLine(input, max));

    [Fact]
    public void Window_returns_a_continuation_hint()
    {
        var text = new string('x', 100);
        var first = Fmt.Window(text, 0, 40, "body_offset");
        Assert.Contains("60 more chars; body_offset=40", first);
        var rest = Fmt.Window(text, 40, 100, "body_offset");
        Assert.DoesNotContain("more chars", rest);
        Assert.Equal(60, rest.Length);
        Assert.Equal("(end of body)", Fmt.Window(text, 100, 10, "body_offset"));
        Assert.Equal("", Fmt.Window("", 0, 10, "body_offset"));
    }

    [Theory]
    [InlineData(500, "500B")]
    [InlineData(2048, "2KB")]
    [InlineData(5 * 1048576, "5MB")]
    public void Size_formats_bytes(long bytes, string expected) => Assert.Equal(expected, Fmt.Size(bytes));

    [Fact]
    public void Flags_are_compact_and_empty_when_none()
    {
        Assert.Equal("", Fmt.Flags());
        Assert.Equal(" [U!@F]", Fmt.Flags(unread: true, high: true, attach: true, flagged: true));
        Assert.Equal(" [R]", Fmt.Flags(recurring: true));
    }

    [Fact]
    public void StripHtml_removes_tags_styles_and_decodes_entities()
    {
        var html = "<html><head><style>p{color:red}</style></head><body><p>Hello&nbsp;<b>World</b> &amp; co</p><br>Line2</body></html>";
        var text = Fmt.StripHtml(html);
        Assert.Contains("World & co", text);
        Assert.Contains("Line2", text);
        Assert.DoesNotContain("color:red", text);
        Assert.DoesNotContain("<", text);
    }

    [Fact]
    public void Dates_use_short_formats()
    {
        var d = new DateTime(DateTime.Now.Year, 10, 22, 14, 5, 0);
        Assert.Equal("10-22 14:05", Fmt.Date(d));
        Assert.Equal($"{DateTime.Now.Year}-10-22T14:05", Fmt.Iso(d));
        Assert.Equal("2020-01-05 10:00", Fmt.Date(new DateTime(2020, 1, 5, 10, 0, 0)));
    }
}
