using System.ComponentModel;
using System.Security.Claims;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol.Server;
using static ExchangeEwsMcp.Exchange.EwsHelpers;

namespace ExchangeEwsMcp.Tools;

[McpServerToolType]
public sealed class PeopleTools(EwsContext ews)
{
    [McpServerTool(Name = "whoami", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Returns the signed-in user, the server time zone, the current local time, whether sending mail is enabled and the access level (ReadOnly, NoDelete, Full).")]
    public async Task<UserInfo> WhoAmI()
    {
        var smtp = await ews.GetCurrentUserSmtpAsync();
        return new UserInfo(
            ews.User.FindFirstValue(ClaimTypes.Name),
            smtp,
            ews.User.FindFirstValue(ClaimTypes.Upn),
            ews.TimeZone.Id,
            ews.Format(ews.Now),
            ews.Options.AllowSend,
            ews.Options.AccessLevel.ToString());
    }

    [McpServerTool(Name = "resolve_names", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Looks up people, distribution lists and rooms in the user's contacts and the global address list by (partial) name or e-mail address.")]
    public async Task<IReadOnlyList<ResolvedName>> ResolveNames(
        [Description("Name, partial name or e-mail address.")] string query)
    {
        var results = await ews.Service.ResolveName(query, ResolveNameSearchLocation.ContactsThenDirectory, true);
        return results.Select(r =>
        {
            var c = r.Contact;
            return new ResolvedName(
                NullIfEmpty(r.Mailbox.Name),
                NullIfEmpty(r.Mailbox.Address),
                r.Mailbox.MailboxType?.ToString(),
                NullIfEmpty(Try(() => c?.Department)),
                NullIfEmpty(Try(() => c?.JobTitle)),
                NullIfEmpty(Try(() => c?.OfficeLocation)),
                NullIfEmpty(Try(() => c is not null && c.PhoneNumbers.TryGetValue(PhoneNumberKey.BusinessPhone, out var p) ? p : null)),
                r.Mailbox.MailboxType == MailboxType.Contact ? "contacts" : "directory");
        }).ToList();
    }

    [McpServerTool(Name = "expand_group", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("""
        Lists the members of a distribution list (e.g. to answer "Am I in list X?" or "Who is in X?").
        Exchange cannot list the groups a person is a member of; check specific lists instead.
        """)]
    public async Task<GroupExpansion> ExpandGroup(
        [Description("Distribution list: e-mail address, alias or name.")] string group,
        [Description("Also expand nested distribution lists. Default false.")] bool recursive = false)
    {
        var address = (await ResolveOneAsync(ews.Service, group.Trim(), ResolveNameSearchLocation.DirectoryOnly, "Distribution list",
            m => m.MailboxType == MailboxType.PublicGroup)).Address;
        var members = new List<ListMember>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var expanded = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { address };
        var pending = new Queue<string>([address]);
        var complete = true;

        // shortcut: stops after 2000 members, enough for "am I in it?"; page through nested lists if larger ones matter.
        while (pending.TryDequeue(out var current) && members.Count < 2000)
        {
            var result = await ews.Service.ExpandGroup(current);
            complete &= result.IncludesAllMembers;
            foreach (var m in result.Members)
            {
                var isGroup = m.MailboxType is MailboxType.PublicGroup or MailboxType.ContactGroup;
                if (!seen.Add(m.Address ?? m.Name ?? "")) continue;
                members.Add(new ListMember(NullIfEmpty(m.Name), NullIfEmpty(m.Address), isGroup ? "group" : "person",
                    current == address ? null : current));
                if (recursive && isGroup && !string.IsNullOrEmpty(m.Address) && expanded.Add(m.Address))
                {
                    pending.Enqueue(m.Address);
                }
            }
        }
        return new GroupExpansion(address, members.Count, complete && pending.Count == 0, members);
    }
}
