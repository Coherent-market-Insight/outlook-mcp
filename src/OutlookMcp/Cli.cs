using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using OutlookMcp.Com;
using OutlookMcp.Tools;

namespace OutlookMcp;

/// <summary>Everything that is not "serve": doctor, installers, tool listing.</summary>
internal static class Cli
{
    public const string Instructions =
        "Outlook desktop (COM). Ids are opaque: copy them from list/search results. " +
        "mail_list returns one line per message; call mail_get only for messages you need to read. " +
        "Sending is explicit: mail_send/mail_draft_send send; mail_reply and mail_forward only save drafts unless send=true. " +
        "Call status once to learn account addresses; pass `account` to use a non-default mailbox. " +
        "Dates are local time, ISO 8601.";

    public const string Help = """
        outlook-mcp - Outlook MCP server for Windows (Coherent Market Insights)

        Usage:
          outlook-mcp [serve] [flags]        Run the MCP server on stdio (default)
          outlook-mcp doctor                 Check Outlook, accounts and speed
          outlook-mcp install --client <c>   Register in an MCP client (see below)
          outlook-mcp install-skill          Install the Outlook skill for Claude Code
          outlook-mcp tools [--markdown]     List tools exposed under the current flags
          outlook-mcp version

        Serve flags (or the matching OUTLOOK_MCP_* environment variables):
          --read-only                   Expose only read tools
          --disable-send                Hide every tool that can send mail
          --toolsets mail,calendar      Limit toolsets: mail, calendar, contacts, tasks
          --disable-tools a,b           Hide specific tools by name
          --allowed-domains x.com,y.com Refuse to send to any other domain

        Install clients: claude-code, claude-desktop, cursor, vscode, gemini, opencode, codex
          --client <c>        Required
          --config-dir <dir>  claude-code only: profile dir holding .claude.json (default ~)
          --exe <path>        Register this executable instead of the running one
          --               Everything after is passed to the server, e.g. -- --read-only
        """;

    // ------------------------------------------------------------------ doctor

    public static async Task<int> Doctor(string version)
    {
        var ok = true;
        void Line(bool pass, string text) { Console.WriteLine($"{(pass ? "ok  " : "FAIL")} {text}"); ok &= pass; }

        Console.WriteLine($"outlook-mcp {version}");
        Line(OperatingSystem.IsWindows(), $"Windows {Environment.OSVersion.Version}");
        var progId = Type.GetTypeFromProgID("Outlook.Application");
        Line(progId is not null, progId is null ? "Outlook COM not registered (install classic Outlook)" : "Outlook COM registered");
        if (progId is null) return 1;
        try
        {
            var sw = Stopwatch.StartNew();
            var text = await Sta.Run(SystemTools.Describe);
            Line(true, $"connected in {sw.ElapsedMilliseconds} ms");
            Console.WriteLine(text);
            sw.Restart();
            var inbox = await Sta.Run(() => (int)Ol.DefaultFolder(Ol.Inbox, null).Items.Count);
            Line(true, $"inbox has {inbox} items (read in {sw.ElapsedMilliseconds} ms)");
        }
        catch (Exception ex) { Line(false, ex.Message); }
        return ok ? 0 : 1;
    }

    // ------------------------------------------------------------------ tools listing

    public static int ListTools(string[] args)
    {
        Current.Settings = Settings.Load(args);
        var entries = Registry.All().Where(e => Current.Settings.Allows(e.Meta)).OrderBy(e => e.Meta.Toolset).ThenBy(e => e.Meta.Name).ToList();
        if (args.Contains("--markdown"))
        {
            foreach (var group in entries.GroupBy(e => e.Meta.Toolset))
            {
                Console.WriteLine($"### {group.Key}\n");
                Console.WriteLine("| Tool | Access | What it does |\n|---|---|---|");
                foreach (var e in group)
                    Console.WriteLine($"| `{e.Meta.Name}` | {e.Meta.Access.ToString().ToLowerInvariant()} | {e.Method.GetCustomAttribute<DescriptionAttribute>()?.Description.Replace("|", "\\|")} |");
                Console.WriteLine();
            }
        }
        else
            foreach (var e in entries) Console.WriteLine($"{e.Meta.Name,-24} {e.Meta.Toolset,-9} {e.Meta.Access}");
        return 0;
    }

    // ------------------------------------------------------------------ install into MCP clients

    // OUTLOOK_MCP_HOME redirects every client config path; the installer tests use it so they never touch real configs.
    private static string Home => Environment.GetEnvironmentVariable("OUTLOOK_MCP_HOME") ?? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    private static string AppData => Environment.GetEnvironmentVariable("OUTLOOK_MCP_HOME") is { } h
        ? Path.Combine(h, "AppData", "Roaming")
        : Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

    public static int Install(string[] args)
    {
        string? Flag(string n) => args.SkipWhile(a => a != n).Skip(1).FirstOrDefault();
        var client = Flag("--client")?.ToLowerInvariant();
        if (client is null) { Console.Error.WriteLine("Missing --client.\n\n" + Help); return 2; }

        var exe = Flag("--exe") ?? Environment.ProcessPath ?? throw new InvalidOperationException("Cannot determine executable path.");
        var sep = Array.IndexOf(args, "--");
        var extra = sep >= 0 ? args[(sep + 1)..] : [];

        try
        {
            switch (client)
            {
                case "claude-code":
                    var dir = Flag("--config-dir") ?? Home;
                    var file = Path.Combine(dir, ".claude.json");
                    if (!File.Exists(file) && Path.GetFileName(dir.TrimEnd('\\', '/')).StartsWith(".claude")) file = Path.Combine(dir, ".claude.json");
                    EditJson(file, ["mcpServers"], StdioEntry(exe, extra, withType: true));
                    break;
                case "claude-desktop":
                    EditJson(Path.Combine(AppData, "Claude", "claude_desktop_config.json"), ["mcpServers"], StdioEntry(exe, extra, withType: false));
                    break;
                case "cursor":
                    EditJson(Path.Combine(Home, ".cursor", "mcp.json"), ["mcpServers"], StdioEntry(exe, extra, withType: false));
                    break;
                case "vscode":
                    EditJson(Path.Combine(AppData, "Code", "User", "mcp.json"), ["servers"], StdioEntry(exe, extra, withType: true));
                    break;
                case "gemini":
                    EditJson(Path.Combine(Home, ".gemini", "settings.json"), ["mcpServers"], StdioEntry(exe, extra, withType: false));
                    break;
                case "opencode":
                    var oc = new JsonObject
                    {
                        ["type"] = "local",
                        ["command"] = new JsonArray(new[] { exe }.Concat(extra).Select(s => (JsonNode)s!).ToArray()),
                        ["enabled"] = true,
                    };
                    EditJson(Path.Combine(Home, ".config", "opencode", "opencode.json"), ["mcp"], oc);
                    break;
                case "codex":
                    InstallCodex(exe, extra);
                    break;
                default:
                    Console.Error.WriteLine($"Unknown client '{client}'.");
                    return 2;
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Install failed: {ex.Message}");
            return 1;
        }
    }

    private static JsonObject StdioEntry(string exe, string[] extra, bool withType)
    {
        var o = new JsonObject();
        if (withType) o["type"] = "stdio";
        o["command"] = exe;
        o["args"] = new JsonArray(extra.Select(s => (JsonNode)s!).ToArray());
        return o;
    }

    private static void EditJson(string path, string[] keyPath, JsonObject entry)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        JsonObject root;
        if (File.Exists(path))
        {
            var backup = path + ".bak-outlook-mcp";
            File.Copy(path, backup, overwrite: true);
            var text = File.ReadAllText(path);
            root = string.IsNullOrWhiteSpace(text)
                ? new JsonObject()
                : JsonNode.Parse(text, documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true })?.AsObject()
                  ?? new JsonObject();
        }
        else root = new JsonObject();

        var cursor = root;
        foreach (var key in keyPath)
        {
            if (cursor[key] is not JsonObject next) { next = new JsonObject(); cursor[key] = next; }
            cursor = next;
        }
        var existed = cursor.ContainsKey("outlook");
        cursor["outlook"] = entry;
        File.WriteAllText(path, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"{(existed ? "updated" : "added")} 'outlook' in {path}" + (File.Exists(path + ".bak-outlook-mcp") ? " (backup: .bak-outlook-mcp)" : ""));
    }

    private static void InstallCodex(string exe, string[] extra)
    {
        var path = Path.Combine(Home, ".codex", "config.toml");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var text = File.Exists(path) ? File.ReadAllText(path) : "";
        if (text.Contains("[mcp_servers.outlook]", StringComparison.Ordinal))
        {
            Console.WriteLine($"'outlook' is already in {path}; edit it by hand to change flags.");
            return;
        }
        if (File.Exists(path)) File.Copy(path, path + ".bak-outlook-mcp", overwrite: true);
        var sb = new StringBuilder(text.TrimEnd());
        sb.Append("\n\n[mcp_servers.outlook]\n");
        sb.Append($"command = '{exe}'\n");
        sb.Append($"args = [{string.Join(", ", extra.Select(a => $"\"{a}\""))}]\n");
        File.WriteAllText(path, sb.ToString());
        Console.WriteLine($"added 'outlook' to {path}");
    }

    // ------------------------------------------------------------------ install skill

    public static int InstallSkill(string[] args)
    {
        string? Flag(string n) => args.SkipWhile(a => a != n).Skip(1).FirstOrDefault();
        var root = Flag("--dir") ?? Path.Combine(Flag("--config-dir") ?? Path.Combine(Home, ".claude"), "skills");
        var target = Path.Combine(root, "outlook");
        using var stream = typeof(Cli).Assembly.GetManifestResourceStream("skill/SKILL.md")
            ?? throw new InvalidOperationException("Skill resource missing from this build.");
        Directory.CreateDirectory(target);
        using var file = File.Create(Path.Combine(target, "SKILL.md"));
        stream.CopyTo(file);
        Console.WriteLine($"skill installed: {Path.Combine(target, "SKILL.md")}");
        return 0;
    }
}
