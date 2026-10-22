using System.ComponentModel;
using System.Reflection;
using System.Text.RegularExpressions;
using OutlookMcp;

namespace OutlookMcp.Tests;

/// <summary>Quality gates on the tool surface: naming, descriptions, access levels and the token budget.</summary>
public class RegistryTests
{
    private static readonly List<Registry.Entry> All = Registry.All().ToList();

    [Fact]
    public void Tool_count_is_stable() => Assert.Equal(31, All.Count);

    [Fact]
    public void Names_are_unique_snake_case()
    {
        Assert.Equal(All.Count, All.Select(e => e.Meta.Name).Distinct().Count());
        Assert.All(All, e => Assert.Matches("^[a-z]+(_[a-z]+)*$", e.Meta.Name));
    }

    [Fact]
    public void Every_tool_has_a_description_and_every_parameter_is_documented()
    {
        foreach (var e in All)
        {
            var d = e.Method.GetCustomAttribute<DescriptionAttribute>()?.Description;
            Assert.False(string.IsNullOrWhiteSpace(d), $"{e.Meta.Name} has no description");
            foreach (var p in e.Method.GetParameters())
                Assert.True(p.GetCustomAttribute<DescriptionAttribute>() is { Description.Length: > 0 },
                    $"{e.Meta.Name}.{p.Name} has no description");
        }
    }

    [Fact]
    public void Tools_belong_to_a_known_toolset()
        => Assert.All(All, e => Assert.Contains(e.Meta.Toolset, Settings.AllToolsets));

    [Fact]
    public void Sending_tools_are_marked_Send_and_deleting_tools_Delete()
    {
        foreach (var name in new[] { "mail_send", "mail_draft_send", "mail_reply", "mail_forward", "calendar_respond" })
            Assert.Equal(Access.Send, All.Single(e => e.Meta.Name == name).Meta.Access);
        foreach (var name in new[] { "mail_delete", "calendar_event_delete", "contact_delete", "task_delete" })
            Assert.Equal(Access.Delete, All.Single(e => e.Meta.Name == name).Meta.Access);
    }

    [Fact]
    public void Read_only_mode_exposes_only_read_tools_and_includes_the_basics()
    {
        var names = Registry.Build(Settings.Load(["--read-only"])).Select(t => t.ProtocolTool.Name).ToList();
        Assert.Contains("mail_list", names);
        Assert.Contains("calendar_events", names);
        Assert.Contains("status", names);
        Assert.DoesNotContain("mail_send", names);
        Assert.DoesNotContain("mail_draft_save", names);
        Assert.DoesNotContain("task_save", names);
        Assert.Equal(All.Count(e => e.Meta.Access == Access.Read), names.Count);
    }

    [Fact]
    public void Annotations_flag_read_only_and_destructive_tools()
    {
        var tools = Registry.Build(new Settings()).ToDictionary(t => t.ProtocolTool.Name);
        Assert.True(tools["mail_list"].ProtocolTool.Annotations?.ReadOnlyHint);
        Assert.False(tools["mail_send"].ProtocolTool.Annotations?.ReadOnlyHint ?? false);
        Assert.True(tools["mail_delete"].ProtocolTool.Annotations?.DestructiveHint);
    }

    [Fact]
    public void Full_tool_surface_stays_within_the_token_budget()
    {
        // ~4 characters per token. This guards against descriptions creeping up over time.
        var chars = Registry.Build(new Settings()).Sum(t => t.ProtocolTool.InputSchema.GetRawText().Length + (t.ProtocolTool.Description?.Length ?? 0) + t.ProtocolTool.Name.Length);
        Assert.True(chars < 30_000, $"tool definitions are {chars} chars (~{chars / 4} tokens); budget is 30000 chars");
    }

    [Fact]
    public void Read_only_surface_is_much_smaller_than_the_full_one()
    {
        int Size(Settings s) => Registry.Build(s).Sum(t => t.ProtocolTool.InputSchema.GetRawText().Length);
        Assert.True(Size(Settings.Load(["--read-only"])) < Size(new Settings()) * 0.6);
    }

    [Fact]
    public void Descriptions_do_not_use_em_dashes()
        => Assert.All(All, e => Assert.DoesNotContain('—', e.Method.GetCustomAttribute<DescriptionAttribute>()!.Description));

    [Fact]
    public void Server_instructions_stay_short() => Assert.True(Cli.Instructions.Length < 700);
}
