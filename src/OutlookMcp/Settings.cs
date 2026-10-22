namespace OutlookMcp;

/// <summary>What a tool is allowed to do. Drives MCP annotations and the --read-only gate.</summary>
internal enum Access { Read, Write, Send, Delete }

[AttributeUsage(AttributeTargets.Method)]
internal sealed class ToolMetaAttribute(string name, string toolset, Access access) : Attribute
{
    public string Name { get; } = name;
    public string Toolset { get; } = toolset;
    public Access Access { get; } = access;
}

/// <summary>Runtime configuration, from environment variables and command-line flags.</summary>
internal sealed class Settings
{
    public static readonly string[] AllToolsets = ["system", "mail", "calendar", "contacts", "tasks"];

    public bool ReadOnly { get; init; }
    public bool DisableSend { get; init; }
    public HashSet<string> Toolsets { get; init; } = new(AllToolsets, StringComparer.OrdinalIgnoreCase);
    public HashSet<string> DisabledTools { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> AllowedDomains { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public int TimeoutSeconds { get; init; } = 60;
    public int BodyMaxChars { get; init; } = 3000;
    public bool Debug { get; init; }
    public string? LogFile { get; init; }

    public static Settings Load(string[] args)
    {
        string? Flag(string name) => args.SkipWhile(a => a != name).Skip(1).FirstOrDefault();
        bool Has(string name) => args.Contains(name);
        static bool Truthy(string? v) => v is "1" or "true" or "TRUE" or "True" or "yes";
        static HashSet<string> Csv(string? v) =>
            new((v ?? "").Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
                StringComparer.OrdinalIgnoreCase);

        static string? env(string n) => Environment.GetEnvironmentVariable(n);
        var toolsets = Csv(Flag("--toolsets") ?? env("OUTLOOK_MCP_TOOLSETS"));
        if (toolsets.Count == 0) toolsets = new(AllToolsets, StringComparer.OrdinalIgnoreCase);
        toolsets.Add("system");

        return new Settings
        {
            ReadOnly = Has("--read-only") || Truthy(env("OUTLOOK_MCP_READ_ONLY")),
            DisableSend = Has("--disable-send") || Truthy(env("OUTLOOK_MCP_DISABLE_SEND")),
            Toolsets = toolsets,
            DisabledTools = Csv(Flag("--disable-tools") ?? env("OUTLOOK_MCP_DISABLED_TOOLS")),
            AllowedDomains = Csv(Flag("--allowed-domains") ?? env("OUTLOOK_MCP_ALLOWED_DOMAINS")),
            TimeoutSeconds = int.TryParse(env("OUTLOOK_MCP_TIMEOUT_SECONDS"), out var t) && t > 0 ? t : 60,
            Debug = Has("--debug") || Truthy(env("OUTLOOK_MCP_DEBUG")) || Flag("--log-file") is not null || env("OUTLOOK_MCP_LOG_FILE") is not null,
            LogFile = Flag("--log-file") ?? env("OUTLOOK_MCP_LOG_FILE"),
            BodyMaxChars = int.TryParse(env("OUTLOOK_MCP_BODY_MAX_CHARS"), out var b) && b > 200 ? b : 3000,
        };
    }

    public bool Allows(ToolMetaAttribute meta)
    {
        if (!Toolsets.Contains(meta.Toolset)) return false;
        if (DisabledTools.Contains(meta.Name)) return false;
        if (ReadOnly && meta.Access != Access.Read) return false;
        if (DisableSend && meta.Access == Access.Send) return false;
        return true;
    }
}

internal static class Current
{
    public static Settings Settings { get; set; } = new();
}

/// <summary>An error whose message is safe and useful to show the model verbatim.</summary>
internal sealed class ToolError(string message) : Exception(message);
