using OutlookMcp;
using OutlookMcp.Com;

namespace OutlookMcp.Tests;

public class SettingsTests
{
    private static ToolMetaAttribute Meta(string name, string set, Access a) => new(name, set, a);

    [Fact]
    public void Defaults_allow_everything()
    {
        var s = Settings.Load([]);
        Assert.False(s.ReadOnly);
        Assert.False(s.DisableSend);
        Assert.False(s.Debug);
        foreach (var a in Enum.GetValues<Access>())
            Assert.True(s.Allows(Meta("x", "mail", a)));
    }

    [Fact]
    public void ReadOnly_allows_only_read_tools()
    {
        var s = Settings.Load(["--read-only"]);
        Assert.True(s.Allows(Meta("a", "mail", Access.Read)));
        Assert.False(s.Allows(Meta("b", "mail", Access.Write)));
        Assert.False(s.Allows(Meta("c", "mail", Access.Send)));
        Assert.False(s.Allows(Meta("d", "mail", Access.Delete)));
    }

    [Fact]
    public void DisableSend_hides_only_send_tools()
    {
        var s = Settings.Load(["--disable-send"]);
        Assert.False(s.Allows(Meta("a", "mail", Access.Send)));
        Assert.True(s.Allows(Meta("b", "mail", Access.Write)));
        Assert.True(s.Allows(Meta("c", "mail", Access.Delete)));
    }

    [Fact]
    public void Toolsets_limit_tools_but_always_keep_system()
    {
        var s = Settings.Load(["--toolsets", "calendar,tasks"]);
        Assert.True(s.Allows(Meta("a", "calendar", Access.Read)));
        Assert.True(s.Allows(Meta("b", "tasks", Access.Write)));
        Assert.True(s.Allows(Meta("status", "system", Access.Read)));
        Assert.False(s.Allows(Meta("c", "mail", Access.Read)));
        Assert.False(s.Allows(Meta("d", "contacts", Access.Read)));
    }

    [Fact]
    public void DisabledTools_are_hidden_case_insensitively()
    {
        var s = Settings.Load(["--disable-tools", "Mail_Delete, calendar_event_delete"]);
        Assert.False(s.Allows(Meta("mail_delete", "mail", Access.Delete)));
        Assert.False(s.Allows(Meta("calendar_event_delete", "calendar", Access.Delete)));
        Assert.True(s.Allows(Meta("mail_list", "mail", Access.Read)));
    }

    [Fact]
    public void AllowedDomains_parse_into_a_set()
    {
        var s = Settings.Load(["--allowed-domains", "Example.com;test.org"]);
        Assert.Contains("example.com", s.AllowedDomains);
        Assert.Contains("TEST.ORG", s.AllowedDomains);
        Assert.Equal(2, s.AllowedDomains.Count);
    }

    [Fact]
    public void Debug_turns_on_with_flag_or_log_file()
    {
        Assert.True(Settings.Load(["--debug"]).Debug);
        var s = Settings.Load(["--log-file", "x.log"]);
        Assert.True(s.Debug);
        Assert.Equal("x.log", s.LogFile);
    }

    [Fact]
    public void Environment_variables_are_honoured()
    {
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_READ_ONLY", "1");
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_TIMEOUT_SECONDS", "15");
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_BODY_MAX_CHARS", "1200");
        try
        {
            var s = Settings.Load([]);
            Assert.True(s.ReadOnly);
            Assert.Equal(15, s.TimeoutSeconds);
            Assert.Equal(1200, s.BodyMaxChars);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OUTLOOK_MCP_READ_ONLY", null);
            Environment.SetEnvironmentVariable("OUTLOOK_MCP_TIMEOUT_SECONDS", null);
            Environment.SetEnvironmentVariable("OUTLOOK_MCP_BODY_MAX_CHARS", null);
        }
    }

    [Fact]
    public void Invalid_numbers_fall_back_to_defaults()
    {
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_TIMEOUT_SECONDS", "abc");
        try { Assert.Equal(60, Settings.Load([]).TimeoutSeconds); }
        finally { Environment.SetEnvironmentVariable("OUTLOOK_MCP_TIMEOUT_SECONDS", null); }
    }
}

public class OlHelperTests
{
    [Fact]
    public void Alias_is_stable_six_hex_characters()
    {
        var a = Ol.Alias("0000000038A1BB1005E5101AA1BB08002B2A56C20000");
        Assert.Equal(6, a.Length);
        Assert.Matches("^[0-9a-f]{6}$", a);
        Assert.Equal(a, Ol.Alias("0000000038A1BB1005E5101AA1BB08002B2A56C20000"));
        Assert.NotEqual(a, Ol.Alias("another store id"));
    }

    [Theory]
    [InlineData("2026-10-30T14:00", 2026, 10, 30, 14, 0)]
    [InlineData("2026-10-30", 2026, 10, 30, 0, 0)]
    public void ParseDate_accepts_iso_formats(string input, int y, int mo, int d, int h, int mi)
        => Assert.Equal(new DateTime(y, mo, d, h, mi, 0), Ol.ParseDate(input, "start"));

    [Fact]
    public void ParseDate_gives_an_actionable_error()
    {
        var ex = Assert.Throws<ToolError>(() => Ol.ParseDate("next tuesday-ish", "since"));
        Assert.Contains("'since'", ex.Message);
        Assert.Contains("ISO 8601", ex.Message);
    }

    [Fact]
    public void JetDate_is_locale_independent()
        => Assert.Equal("10/30/2026 02:05 PM", Ol.JetDate(new DateTime(2026, 10, 30, 14, 5, 0)));

    [Fact]
    public void Trunc_caps_length() => Assert.Equal("abc", Ol.Trunc("abcdef", 3));
}
