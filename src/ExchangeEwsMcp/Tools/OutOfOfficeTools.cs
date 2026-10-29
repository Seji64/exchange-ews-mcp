using System.ComponentModel;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace ExchangeEwsMcp.Tools;

[McpServerToolType]
public sealed class OutOfOfficeTools(EwsContext ews)
{
    [McpServerTool(Name = "get_out_of_office", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Gets the automatic replies (out of office) settings.")]
    public async Task<OutOfOfficeInfo> GetOutOfOffice([Description(MailTools.MailboxDescription)] string? mailbox = null)
    {
        var address = string.IsNullOrWhiteSpace(mailbox) ? await ews.GetCurrentUserSmtpAsync() : mailbox.Trim();
        var oof = await ews.Service.GetUserOofSettings(address);
        return new OutOfOfficeInfo(
            address,
            oof.State.ToString(),
            oof.State == OofState.Scheduled ? ews.Format(oof.Duration.StartTime) : null,
            oof.State == OofState.Scheduled ? ews.Format(oof.Duration.EndTime) : null,
            oof.ExternalAudience.ToString(),
            oof.InternalReply?.Message,
            oof.ExternalReply?.Message);
    }

    [McpServerTool(Name = "set_out_of_office", Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("""
        Configures automatic replies (out of office). Use state 'scheduled' with start/end for a time range,
        'enabled' to turn on now, 'disabled' to turn off. Replies may contain HTML.
        """)]
    public async Task<OutOfOfficeInfo> SetOutOfOffice(
        [Description("'enabled', 'scheduled' or 'disabled'.")] string state,
        [Description("Reply sent to people inside the organization. Omit to keep the current text.")] string? internalReply = null,
        [Description("Reply sent to external senders. Omit to keep the current text.")] string? externalReply = null,
        [Description("Start (ISO 8601), required for 'scheduled'.")] string? start = null,
        [Description("End (ISO 8601), required for 'scheduled'.")] string? end = null,
        [Description("External audience: 'none', 'known' (contacts only) or 'all'. Omit to keep the current setting.")] string? externalAudience = null,
        [Description(MailTools.MailboxDescription)] string? mailbox = null)
    {
        var address = string.IsNullOrWhiteSpace(mailbox) ? await ews.GetCurrentUserSmtpAsync() : mailbox.Trim();
        var oof = await ews.Service.GetUserOofSettings(address);

        oof.State = state.ToLowerInvariant() switch
        {
            "enabled" or "on" => OofState.Enabled,
            "scheduled" => OofState.Scheduled,
            "disabled" or "off" => OofState.Disabled,
            _ => throw new McpException("state must be 'enabled', 'scheduled' or 'disabled'."),
        };
        if (oof.State == OofState.Scheduled)
        {
            if (start is null || end is null) throw new McpException("'start' and 'end' are required for state 'scheduled'.");
            var from = ews.ParseDateTime(start, nameof(start));
            var to = ews.ParseDateTime(end, nameof(end));
            if (to <= from) throw new McpException("'end' must be after 'start'.");
            oof.Duration = new TimeWindow(ToServiceTime(from), ToServiceTime(to));
        }
        if (internalReply is not null) oof.InternalReply = new OofReply(internalReply);
        if (externalReply is not null) oof.ExternalReply = new OofReply(externalReply);
        if (externalAudience is not null)
        {
            oof.ExternalAudience = externalAudience.ToLowerInvariant() switch
            {
                "none" => OofExternalAudience.None,
                "known" => OofExternalAudience.Known,
                "all" => OofExternalAudience.All,
                _ => throw new McpException("externalAudience must be 'none', 'known' or 'all'."),
            };
        }

        await ews.Service.SetUserOofSettings(address, oof);
        return await GetOutOfOffice(mailbox);
    }

    // OOF durations are interpreted in the service time zone; convert absolute (UTC) values accordingly.
    private DateTime ToServiceTime(DateTime value) =>
        value.Kind == DateTimeKind.Utc ? TimeZoneInfo.ConvertTimeFromUtc(value, ews.TimeZone) : value;
}
