using System.ComponentModel;
using System.Text;
using OutlookMcp.Com;

namespace OutlookMcp.Tools;

internal static class ContactTools
{
    [ToolMeta("contacts_search", "contacts", Access.Read)]
    [Description("Search saved contacts by name, email or company. One line each: id, name, company, title, email, phone.")]
    public static Task<string> ContactsSearch(
        [Description("Text to find; omit to list")] string? query = null,
        [Description("Account whose contacts to search; default account")] string? account = null,
        [Description("Max rows (1-100)")] int limit = 15,
        [Description("Rows to skip")] int offset = 0)
        => Sta.Run<string>(() =>
        {
            limit = Math.Clamp(limit, 1, 100);
            dynamic folder = Ol.DefaultFolder(Ol.Contacts, account);
            dynamic items = folder.Items;
            if (!string.IsNullOrWhiteSpace(query))
            {
                var q = query!.Replace("'", "''");
                items = items.Restrict("@SQL=(\"urn:schemas:contacts:cn\" LIKE '%" + q + "%' OR \"urn:schemas:contacts:o\" LIKE '%" + q +
                                       "%' OR \"urn:schemas:contacts:email1\" LIKE '%" + q + "%' OR \"urn:schemas:contacts:email2\" LIKE '%" + q + "%')");
            }
            items.Sort("[FullName]");
            int total = items.Count;
            var sb = new StringBuilder();
            var shown = 0; var skipped = 0; var more = false;
            dynamic? it = items.GetFirst();
            while (it is not null)
            {
                if (Ol.Safe(() => (int)it.Class, 40) == 40)
                {
                    if (skipped < Math.Max(0, offset)) skipped++;
                    else if (shown >= limit) { more = true; Ol.Release(it); break; }
                    else { sb.AppendLine(ContactLine(it)); shown++; }
                }
                Ol.Release(it);
                it = items.GetNext();
            }
            var head = $"contacts: {shown} shown, {total} match{(more ? $", more: offset={offset + shown}" : "")}";
            return shown == 0 ? head : head + "\n" + sb.ToString().TrimEnd();
        });

    private static string ContactLine(dynamic c)
    {
        var phone = new[] { Ol.S(() => c.MobileTelephoneNumber), Ol.S(() => c.BusinessTelephoneNumber) }.FirstOrDefault(p => p.Length > 0) ?? "";
        return string.Join(" | ", new[]
        {
            Ol.Id(c), Ol.S(() => c.FullName), Ol.S(() => c.CompanyName), Ol.S(() => c.JobTitle), Ol.S(() => c.Email1Address), phone,
        }.Select(s => Fmt.OneLine(s, 50)));
    }

    [ToolMeta("contact_get", "contacts", Access.Read)]
    [Description("All stored fields of one contact.")]
    public static Task<string> ContactGet([Description("Contact id")] string id)
        => Sta.Run<string>(() =>
        {
            dynamic c = Ol.GetItem(id);
            var sb = new StringBuilder();
            void F(string label, string value) { if (value.Length > 0) sb.AppendLine($"{label}: {value}"); }
            F("name", Ol.S(() => c.FullName));
            F("company", Ol.S(() => c.CompanyName));
            F("title", Ol.S(() => c.JobTitle));
            F("department", Ol.S(() => c.Department));
            F("email", Ol.S(() => c.Email1Address));
            F("email2", Ol.S(() => c.Email2Address));
            F("email3", Ol.S(() => c.Email3Address));
            F("business phone", Ol.S(() => c.BusinessTelephoneNumber));
            F("mobile", Ol.S(() => c.MobileTelephoneNumber));
            F("home phone", Ol.S(() => c.HomeTelephoneNumber));
            F("business address", Fmt.OneLine(Ol.S(() => c.BusinessAddress), 200));
            F("web", Ol.S(() => c.WebPage));
            F("categories", Ol.S(() => c.Categories));
            F("notes", Fmt.OneLine(Ol.S(() => c.Body), 600));
            sb.AppendLine($"id: {id}");
            return sb.ToString().TrimEnd();
        });

    [ToolMeta("contact_save", "contacts", Access.Write)]
    [Description("Create a contact, or update one when id is given (only the fields you pass change).")]
    public static Task<string> ContactSave(
        [Description("Existing contact id to update; omit to create")] string? id = null,
        [Description("Full name")] string? full_name = null,
        [Description("Company")] string? company = null,
        [Description("Job title")] string? job_title = null,
        [Description("Primary email")] string? email = null,
        [Description("Second email")] string? email2 = null,
        [Description("Business phone")] string? business_phone = null,
        [Description("Mobile phone")] string? mobile_phone = null,
        [Description("Business address")] string? address = null,
        [Description("Website")] string? web = null,
        [Description("Notes")] string? notes = null,
        [Description("Category names")] string[]? categories = null,
        [Description("Account to create it in; default account")] string? account = null)
        => Sta.Run<string>(() =>
        {
            dynamic c;
            if (id is null)
            {
                if (string.IsNullOrWhiteSpace(full_name) && string.IsNullOrWhiteSpace(email))
                    throw new ToolError("A new contact needs at least full_name or email.");
                c = Ol.DefaultFolder(Ol.Contacts, account).Items.Add(2);
            }
            else c = Ol.GetItem(id);
            if (full_name is not null) c.FullName = full_name;
            if (company is not null) c.CompanyName = company;
            if (job_title is not null) c.JobTitle = job_title;
            if (email is not null) c.Email1Address = email;
            if (email2 is not null) c.Email2Address = email2;
            if (business_phone is not null) c.BusinessTelephoneNumber = business_phone;
            if (mobile_phone is not null) c.MobileTelephoneNumber = mobile_phone;
            if (address is not null) c.BusinessAddress = address;
            if (web is not null) c.WebPage = web;
            if (notes is not null) c.Body = notes;
            if (categories is not null) c.Categories = string.Join(", ", categories);
            c.Save();
            return $"{(id is null ? "created" : "updated")}: {ContactLine(c)}";
        });

    [ToolMeta("contact_delete", "contacts", Access.Delete)]
    [Description("Delete contacts (moved to Deleted Items).")]
    public static Task<string> ContactDelete([Description("Contact ids")] string[] ids)
        => Sta.Run<string>(() =>
        {
            foreach (var id in ids) Ol.GetItem(id).Delete();
            return $"deleted {ids.Length}";
        });

    [ToolMeta("recipient_resolve", "contacts", Access.Read)]
    [Description("Resolve names or partial addresses to email addresses using saved contacts and, on Exchange, the company address list. Use before emailing someone you only know by name.")]
    public static Task<string> RecipientResolve([Description("Names or partial addresses")] string[] names)
        => Sta.Run<string>(() =>
        {
            var sb = new StringBuilder();
            foreach (var n in names)
            {
                dynamic r = Ol.Ns.CreateRecipient(n);
                var ok = (bool)r.Resolve();
                if (!ok) { sb.AppendLine($"{n} → not resolved or ambiguous"); continue; }
                var line = $"{n} → {(string)r.Name} <{Ol.RecipientSmtp(r)}>";
                var eu = Ol.Safe<object?>(() => r.AddressEntry.GetExchangeUser(), null);
                if (eu is not null)
                {
                    dynamic u = eu;
                    var extra = new[] { Ol.S(() => u.JobTitle), Ol.S(() => u.Department), Ol.S(() => u.OfficeLocation) }.Where(s => s.Length > 0);
                    line += " | " + string.Join(", ", extra);
                }
                sb.AppendLine(line.TrimEnd(' ', '|'));
            }
            return sb.ToString().TrimEnd();
        });
}
