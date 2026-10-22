using System.Reflection;
using ModelContextProtocol.Server;

namespace OutlookMcp;

/// <summary>
/// Discovers every [ToolMeta] method and applies the runtime policy (toolsets, read-only,
/// disabled tools). The same registry feeds the server, the docs generator and the tests.
/// </summary>
internal static class Registry
{
    public sealed record Entry(ToolMetaAttribute Meta, MethodInfo Method);

    public static IEnumerable<Entry> All()
    {
        var types = typeof(Registry).Assembly.GetTypes().Where(t => t.Namespace == "OutlookMcp.Tools" && t.IsClass);
        foreach (var type in types)
            foreach (var m in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                var meta = m.GetCustomAttribute<ToolMetaAttribute>();
                if (meta is not null) yield return new Entry(meta, m);
            }
    }

    public static IEnumerable<McpServerTool> Build(Settings settings)
    {
        foreach (var e in All().Where(e => settings.Allows(e.Meta)).OrderBy(e => e.Meta.Name, StringComparer.Ordinal))
        {
            yield return McpServerTool.Create(e.Method, target: null, new McpServerToolCreateOptions
            {
                Name = e.Meta.Name,
                ReadOnly = e.Meta.Access == Access.Read,
                Destructive = e.Meta.Access is Access.Delete or Access.Send,
                Idempotent = e.Meta.Access == Access.Read,
                OpenWorld = e.Meta.Access == Access.Send,
            });
        }
    }
}
