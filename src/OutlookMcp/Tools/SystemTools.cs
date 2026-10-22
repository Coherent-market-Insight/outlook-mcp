using System.ComponentModel;
using System.Text;
using OutlookMcp.Com;

namespace OutlookMcp.Tools;

internal static class SystemTools
{
    [ToolMeta("status", "system", Access.Read)]
    [Description("Outlook version, signed-in accounts (with SMTP address and type), default account and store, and this server's active policy. Call once to learn which accounts exist.")]
    public static Task<string> Status() => Sta.Run(Describe);

    internal static string Describe()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"outlook {Ol.S(() => Ol.App.Version)} | profile {Ol.S(() => Ol.Ns.CurrentProfileName)}");
        string defStore = Ol.S(() => Ol.Ns.DefaultStore.DisplayName);
        sb.AppendLine("accounts (type: 0=Exchange 1=IMAP 2=POP 3=HTTP 4=EAS):");
        foreach (dynamic a in Ol.Ns.Accounts)
        {
            var isDefault = Ol.S(() => a.DeliveryStore.DisplayName) == defStore;
            sb.AppendLine($"  {(string)a.SmtpAddress} | {(string)a.DisplayName} | type {(int)a.AccountType}{(isDefault ? " | default" : "")}");
        }
        var s = Current.Settings;
        sb.Append("policy: ")
          .Append(s.ReadOnly ? "read-only" : "read-write")
          .Append(s.DisableSend ? ", send disabled" : "")
          .Append(s.AllowedDomains.Count > 0 ? $", send only to {string.Join("/", s.AllowedDomains)}" : "")
          .Append($", toolsets {string.Join("/", s.Toolsets)}");
        return sb.ToString();
    }
}
