using System.Text.Json.Nodes;
using OutlookMcp;

namespace OutlookMcp.Tests;

/// <summary>The installer is exercised against a throw-away home directory, never real client configs.</summary>
[Collection("installer")]
public class InstallTests : IDisposable
{
    private readonly string _home = Path.Combine(Path.GetTempPath(), "outlook-mcp-tests-" + Guid.NewGuid().ToString("N"));
    private const string Exe = @"C:\tools\outlook-mcp.exe";

    public InstallTests()
    {
        Directory.CreateDirectory(_home);
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_HOME", _home);
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("OUTLOOK_MCP_HOME", null);
        try { Directory.Delete(_home, true); } catch { }
    }

    private static JsonNode Read(string path) => JsonNode.Parse(File.ReadAllText(path))!;
    private string P(params string[] parts) => Path.Combine([_home, .. parts]);

    private static int Run(params string[] args)
    {
        var sw = new StringWriter(); var old = Console.Out; Console.SetOut(sw);
        try { return Cli.Install(args); } finally { Console.SetOut(old); }
    }

    [Fact]
    public void Claude_code_writes_a_stdio_entry_and_keeps_existing_servers()
    {
        File.WriteAllText(P(".claude.json"), """{"theme":"dark","mcpServers":{"other":{"type":"stdio","command":"x","args":[]}}}""");
        Assert.Equal(0, Run("install", "--client", "claude-code", "--exe", Exe));
        var root = Read(P(".claude.json"));
        Assert.Equal("dark", (string?)root["theme"]);
        Assert.Equal("x", (string?)root["mcpServers"]!["other"]!["command"]);
        Assert.Equal(Exe, (string?)root["mcpServers"]!["outlook"]!["command"]);
        Assert.Equal("stdio", (string?)root["mcpServers"]!["outlook"]!["type"]);
        Assert.True(File.Exists(P(".claude.json.bak-outlook-mcp")));
    }

    [Fact]
    public void Claude_code_honours_config_dir_for_a_second_profile()
    {
        var profile = P(".profile-two");
        Assert.Equal(0, Run("install", "--client", "claude-code", "--config-dir", profile, "--exe", Exe));
        Assert.Equal(Exe, (string?)Read(Path.Combine(profile, ".claude.json"))["mcpServers"]!["outlook"]!["command"]);
    }

    [Fact]
    public void Extra_server_flags_are_passed_through_after_the_separator()
    {
        Assert.Equal(0, Run("install", "--client", "cursor", "--exe", Exe, "--", "--read-only", "--toolsets", "mail"));
        var args = Read(P(".cursor", "mcp.json"))["mcpServers"]!["outlook"]!["args"]!.AsArray().Select(a => (string?)a).ToArray();
        Assert.Equal(["--read-only", "--toolsets", "mail"], args);
    }

    [Fact]
    public void Vscode_uses_the_servers_key()
    {
        Assert.Equal(0, Run("install", "--client", "vscode", "--exe", Exe));
        var root = Read(P("AppData", "Roaming", "Code", "User", "mcp.json"));
        Assert.Equal(Exe, (string?)root["servers"]!["outlook"]!["command"]);
        Assert.Null(root["mcpServers"]);
    }

    [Fact]
    public void Gemini_and_claude_desktop_use_mcpServers()
    {
        Assert.Equal(0, Run("install", "--client", "gemini", "--exe", Exe));
        Assert.Equal(0, Run("install", "--client", "claude-desktop", "--exe", Exe));
        Assert.Equal(Exe, (string?)Read(P(".gemini", "settings.json"))["mcpServers"]!["outlook"]!["command"]);
        Assert.Equal(Exe, (string?)Read(P("AppData", "Roaming", "Claude", "claude_desktop_config.json"))["mcpServers"]!["outlook"]!["command"]);
    }

    [Fact]
    public void Opencode_uses_a_local_command_array()
    {
        Assert.Equal(0, Run("install", "--client", "opencode", "--exe", Exe, "--", "--read-only"));
        var entry = Read(P(".config", "opencode", "opencode.json"))["mcp"]!["outlook"]!;
        Assert.Equal("local", (string?)entry["type"]);
        Assert.Equal([Exe, "--read-only"], entry["command"]!.AsArray().Select(a => (string?)a).ToArray());
        Assert.True((bool)entry["enabled"]!);
    }

    [Fact]
    public void Codex_appends_toml_once()
    {
        Assert.Equal(0, Run("install", "--client", "codex", "--exe", Exe));
        Assert.Equal(0, Run("install", "--client", "codex", "--exe", Exe));
        var toml = File.ReadAllText(P(".codex", "config.toml"));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(toml, @"\[mcp_servers\.outlook\]"));
        Assert.Contains($"command = '{Exe}'", toml);
    }

    [Fact]
    public void Installing_twice_updates_in_place_without_duplicating()
    {
        Assert.Equal(0, Run("install", "--client", "gemini", "--exe", Exe));
        Assert.Equal(0, Run("install", "--client", "gemini", "--exe", @"C:\new\outlook-mcp.exe"));
        var servers = Read(P(".gemini", "settings.json"))["mcpServers"]!.AsObject();
        Assert.Single(servers);
        Assert.Equal(@"C:\new\outlook-mcp.exe", (string?)servers["outlook"]!["command"]);
    }

    [Fact]
    public void Comments_and_trailing_commas_in_existing_json_are_tolerated()
    {
        Directory.CreateDirectory(P(".config", "opencode"));
        File.WriteAllText(P(".config", "opencode", "opencode.json"), "{\n // my config\n \"theme\": \"x\",\n}");
        Assert.Equal(0, Run("install", "--client", "opencode", "--exe", Exe));
        Assert.Equal("x", (string?)Read(P(".config", "opencode", "opencode.json"))["theme"]);
    }

    [Fact]
    public void Unknown_client_and_missing_client_fail_cleanly()
    {
        Assert.Equal(2, Run("install", "--client", "nope", "--exe", Exe));
        Assert.Equal(2, Run("install"));
    }

    [Fact]
    public void Skill_installs_to_the_chosen_directory()
    {
        var dir = P("skills-out");
        var sw = new StringWriter(); var old = Console.Out; Console.SetOut(sw);
        try { Assert.Equal(0, Cli.InstallSkill(["install-skill", "--dir", dir])); } finally { Console.SetOut(old); }
        var text = File.ReadAllText(Path.Combine(dir, "outlook", "SKILL.md"));
        Assert.StartsWith("---", text);
        Assert.Contains("name: outlook", text);
    }
}

[CollectionDefinition("installer", DisableParallelization = true)]
public class InstallerCollection;
