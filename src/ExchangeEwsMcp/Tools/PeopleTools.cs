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
}
