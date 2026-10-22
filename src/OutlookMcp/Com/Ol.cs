using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace OutlookMcp.Com;

/// <summary>Late-bound Outlook COM session: lazy connect, reconnect on RPC failure, id and folder helpers.</summary>
internal static class Ol
{
    // OlDefaultFolders
    public const int Deleted = 3, Outbox = 4, Sent = 5, Inbox = 6, Calendar = 9, Contacts = 10,
                     Journal = 11, Notes = 12, Tasks = 13, Drafts = 16, Junk = 23;

    private static dynamic? _app, _ns;

    public static dynamic App { get { Ensure(); return _app!; } }
    public static dynamic Ns { get { Ensure(); return _ns!; } }

    private static void Ensure()
    {
        if (_app is not null) return;
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new ToolError("Outlook is not installed or not registered for COM. Install Outlook (classic desktop, not 'new Outlook').");
        try { _app = Activator.CreateInstance(type); }
        catch (COMException ex) { throw new ToolError($"Could not start Outlook: {ex.Message.Trim()}"); }
        _ns = _app!.GetNamespace("MAPI");
    }

    private static void Reset()
    {
        _app = null; _ns = null; Stores.Clear();
    }

    /// <summary>Runs <paramref name="func"/>, reconnecting once if Outlook was closed or the RPC link dropped.</summary>
    public static T WithRetry<T>(Func<T> func)
    {
        try { return func(); }
        catch (COMException ex) when (IsDisconnected(ex.HResult))
        {
            Reset();
            return func();
        }
    }

    private static bool IsDisconnected(int hr) =>
        hr is unchecked((int)0x80010108) or unchecked((int)0x800706BA) or unchecked((int)0x800706BE)
           or unchecked((int)0x80010005) or unchecked((int)0x800706BF);

    // ---------- small safe-access helpers ----------

    public static T Safe<T>(Func<T> get, T fallback)
    {
        try { var v = get(); return v is null ? fallback : v; } catch { return fallback; }
    }

    public static void Release(object? o)
    {
        if (o is not null && Marshal.IsComObject(o)) try { Marshal.ReleaseComObject(o); } catch { }
    }

    public static string S(Func<object?> get) => Safe(() => Convert.ToString(get(), CultureInfo.InvariantCulture) ?? "", "");

    // ---------- ids: "<store-alias>:<entry id without the store-wide prefix>" ----------
    //
    // Every EntryID in one store starts with the same 40 hex characters (flags + provider
    // UID). Dropping that prefix cuts an id from ~55 to ~15 characters, which adds up over
    // long lists. If an id ever does not carry the prefix it is emitted whole after a '~'.

    private sealed record StoreInfo(string StoreId, string Prefix);

    private static readonly Dictionary<string, StoreInfo> Stores = new();
    private const int PrefixLength = 40;

    public static string Alias(string storeId)
    {
        var hash = SHA1.HashData(Encoding.UTF8.GetBytes(storeId));
        return Convert.ToHexString(hash, 0, 3).ToLowerInvariant();
    }

    private static string Encode(string storeId, dynamic store, string entry)
    {
        var alias = Alias(storeId);
        if (!Stores.TryGetValue(alias, out var info))
        {
            string root = store.GetRootFolder().EntryID;
            info = new StoreInfo(storeId, root.Length >= PrefixLength ? root[..PrefixLength] : root);
            Stores[alias] = info;
        }
        return entry.StartsWith(info.Prefix, StringComparison.Ordinal)
            ? $"{alias}:{entry[info.Prefix.Length..]}"
            : $"{alias}:~{entry}";
    }

    public static string Id(dynamic item, dynamic? folder = null)
    {
        string entry = item.EntryID;
        dynamic parent = folder ?? item.Parent;
        string storeId = parent.StoreID;
        return Encode(storeId, parent.Store, entry);
    }

    public static string FolderId(dynamic folder)
    {
        string storeId = folder.StoreID;
        return Encode(storeId, folder.Store, (string)folder.EntryID);
    }

    private static (StoreInfo Store, string Entry) Decode(string id)
    {
        var parts = id.Split(':', 2);
        if (parts.Length != 2 || parts[0].Length != 6)
            throw new ToolError($"Bad id '{Trunc(id, 40)}'. Use an id returned by a list or search tool.");
        if (!Stores.TryGetValue(parts[0], out var info))
        {
            foreach (dynamic st in Ns.Stores)
            {
                string sid = st.StoreID;
                string root = st.GetRootFolder().EntryID;
                Stores[Alias(sid)] = new StoreInfo(sid, root.Length >= PrefixLength ? root[..PrefixLength] : root);
            }
            if (!Stores.TryGetValue(parts[0], out info))
                throw new ToolError("Unknown store in id. Call status to refresh.");
        }
        var entry = parts[1].StartsWith('~') ? parts[1][1..] : info.Prefix + parts[1];
        return (info, entry);
    }

    public static dynamic GetItem(string id)
    {
        var (store, entry) = Decode(id);
        try { return Ns.GetItemFromID(entry, store.StoreId); }
        catch (COMException) { throw new ToolError($"Item not found (moved, deleted or bad id): {Trunc(id, 24)}"); }
    }

    public static dynamic GetFolderById(string id)
    {
        var (store, entry) = Decode(id);
        try { return Ns.GetFolderFromID(entry, store.StoreId); }
        catch (COMException) { throw new ToolError("Folder not found."); }
    }

    // ---------- accounts and folders ----------

    public static dynamic? FindAccount(string? spec)
    {
        if (string.IsNullOrWhiteSpace(spec)) return null;
        foreach (dynamic a in Ns.Accounts)
        {
            if (string.Equals((string)a.SmtpAddress, spec, StringComparison.OrdinalIgnoreCase) ||
                string.Equals((string)a.DisplayName, spec, StringComparison.OrdinalIgnoreCase))
                return a;
        }
        throw new ToolError($"No Outlook account matches '{spec}'. Call status to list accounts.");
    }

    public static dynamic DefaultFolder(int kind, string? account)
    {
        var acct = FindAccount(account);
        try { return acct is null ? Ns.GetDefaultFolder(kind) : acct.DeliveryStore.GetDefaultFolder(kind); }
        catch (COMException) { throw new ToolError("That account's store has no such default folder."); }
    }

    private static readonly Dictionary<string, int> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["inbox"] = Inbox, ["sent"] = Sent, ["sent items"] = Sent, ["drafts"] = Drafts, ["deleted"] = Deleted,
        ["deleted items"] = Deleted, ["trash"] = Deleted, ["outbox"] = Outbox, ["junk"] = Junk, ["spam"] = Junk,
        ["calendar"] = Calendar, ["contacts"] = Contacts, ["tasks"] = Tasks, ["notes"] = Notes,
    };

    /// <summary>Resolves "inbox", "Inbox/Clients/Acme", a store name, or a folder id.</summary>
    public static dynamic ResolveFolder(string? spec, string? account, int fallbackKind)
    {
        if (string.IsNullOrWhiteSpace(spec)) return DefaultFolder(fallbackKind, account);
        if (spec.Length >= 8 && spec[6] == ':') return GetFolderById(spec);

        var segments = spec.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        dynamic current;
        int start;
        if (Named.TryGetValue(segments[0], out var kind)) { current = DefaultFolder(kind, account); start = 1; }
        else
        {
            dynamic? root = null;
            foreach (dynamic f in Ns.Folders)
                if (string.Equals((string)f.Name, segments[0], StringComparison.OrdinalIgnoreCase)) { root = f; break; }
            if (root is not null) { current = root; start = 1; }
            else { current = DefaultFolder(Inbox, account).Parent; start = 0; }
        }
        for (var i = start; i < segments.Length; i++)
        {
            dynamic? next = null;
            foreach (dynamic child in current.Folders)
                if (string.Equals((string)child.Name, segments[i], StringComparison.OrdinalIgnoreCase)) { next = child; break; }
            current = next ?? throw new ToolError($"Folder not found: '{segments[i]}' under '{(string)current.Name}'.");
        }
        return current;
    }

    // ---------- recipients ----------

    public static string RecipientSmtp(dynamic recipient)
    {
        try
        {
            string type = recipient.AddressEntry?.Type ?? "";
            if (type == "SMTP") return recipient.Address;
            var smtp = recipient.AddressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E") as string;
            if (!string.IsNullOrEmpty(smtp)) return smtp;
            var eu = recipient.AddressEntry.GetExchangeUser();
            if (eu is not null) return eu.PrimarySmtpAddress;
            return recipient.Address;
        }
        catch { return Safe(() => (string)recipient.Address, ""); }
    }

    // ---------- date helpers ----------

    public static readonly CultureInfo Us = CultureInfo.InvariantCulture;

    public static DateTime ParseDate(string value, string field)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var d)) return d;
        throw new ToolError($"'{field}' is not a valid date/time: '{value}'. Use ISO 8601, e.g. 2026-10-30T14:00.");
    }

    /// <summary>Jet filter date literal.</summary>
    public static string JetDate(DateTime d) => d.ToString("MM/dd/yyyy hh:mm tt", Us);

    public static string Trunc(string s, int max) => s.Length <= max ? s : s[..max];
}
