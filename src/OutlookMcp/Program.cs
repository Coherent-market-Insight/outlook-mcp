using System.Diagnostics;
using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using OutlookMcp;
using OutlookMcp.Com;
using Debug = OutlookMcp.Debug;

var version = typeof(Registry).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion?.Split('+')[0] ?? "1.0.0";
var command = args.Length > 0 && !args[0].StartsWith('-') ? args[0].ToLowerInvariant() : "serve";

switch (command)
{
    case "version":
        Console.WriteLine($"outlook-mcp {version}");
        return 0;
    case "help":
        Console.WriteLine(Cli.Help);
        return 0;
    case "doctor":
        return await Cli.Doctor(version);
    case "install":
        return Cli.Install(args);
    case "install-skill":
        return Cli.InstallSkill(args);
    case "tools":
        return Cli.ListTools(args);
    case "serve":
        break;
    default:
        Console.Error.WriteLine($"Unknown command '{command}'.\n\n{Cli.Help}");
        return 2;
}

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("outlook-mcp needs Windows with the classic Outlook desktop client.");
    return 1;
}

Current.Settings = Settings.Load(args);

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Logging.SetMinimumLevel(LogLevel.Warning);
builder.Logging.AddFilter("ModelContextProtocol", LogLevel.Critical);

builder.Services
    .AddMcpServer(o =>
    {
        o.ServerInfo = new() { Name = "outlook", Title = "Outlook (Coherent Market Insights)", Version = version };
        o.ServerInstructions = Cli.Instructions;
    })
    .WithStdioServerTransport()
    .WithTools(Registry.Build(Current.Settings))
    .WithRequestFilters(f => f.AddCallToolFilter(next => async (ctx, ct) =>
    {
        var name = ctx.Params?.Name ?? "?";
        var sw = Stopwatch.StartNew();
        Debug.Log($"> {name} {Debug.Summarise(ctx.Params?.Arguments)}");
        try
        {
            var result = await next(ctx, ct);
            var chars = result.Content?.OfType<TextContentBlock>().Sum(t => t.Text.Length) ?? 0;
            Debug.Log($"< {name} {(result.IsError == true ? "error" : "ok")} {Debug.Elapsed(sw)} {chars} chars (~{chars / 4} tokens)");
            return result;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var message = ex is ToolError or McpException ? ex.Message : $"{ex.GetType().Name}: {ex.Message}";
            Debug.Log($"! {name} {Debug.Elapsed(sw)} {ex.GetType().Name}: {ex.Message}" + (Debug.Enabled && ex is not ToolError ? "\n" + ex.StackTrace : ""));
            return new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = message }] };
        }
    }));

await builder.Build().RunAsync();
return 0;
