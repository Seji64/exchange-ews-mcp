using System.Globalization;
using System.Security.Claims;
using ExchangeEwsMcp.Auth;
using ExchangeEwsMcp.Configuration;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using ModelContextProtocol;
using Task = System.Threading.Tasks.Task;

namespace ExchangeEwsMcp.Exchange;

/// <summary>Per-request access to EWS on behalf of the authenticated user.</summary>
public sealed class EwsContext(
    IHttpContextAccessor httpContextAccessor,
    ExchangeServiceFactory factory,
    IMemoryCache cache,
    IOptions<ExchangeOptions> options)
{
    private ExchangeService? _service;

    public ExchangeOptions Options => options.Value;

    public TimeZoneInfo TimeZone => factory.TimeZone;

    public ClaimsPrincipal User =>
        httpContextAccessor.HttpContext?.User is { Identity.IsAuthenticated: true } user
            ? user
            : throw new McpException("Not authenticated.");

    public string GrantId => User.FindFirstValue(AccessTokenAuthenticationHandler.GrantIdClaim) ?? throw new McpException("Not authenticated.");

    public ExchangeService Service => _service ??= factory.GetService(GrantId);

    /// <summary>Primary SMTP address of the signed-in user (resolved via EWS, cached).</summary>
    public async Task<string> GetCurrentUserSmtpAsync()
    {
        var key = "smtp:" + GrantId;
        if (cache.TryGetValue(key, out string? cached) && cached is not null) return cached;

        var candidate = User.FindFirstValue(ClaimTypes.Email) ?? User.FindFirstValue(ClaimTypes.Upn)
            ?? throw new McpException("The user's e-mail address is unknown.");
        var smtp = candidate;
        try
        {
            var resolved = await Service.ResolveName(candidate, ResolveNameSearchLocation.DirectoryOnly, false);
            if (resolved.Count == 1 && !string.IsNullOrEmpty(resolved[0].Mailbox.Address))
            {
                smtp = resolved[0].Mailbox.Address;
            }
        }
        catch (ServiceResponseException)
        {
            // Fall back to the claim value.
        }
        cache.Set(key, smtp, TimeSpan.FromHours(4));
        return smtp;
    }

    public FolderId WellKnown(WellKnownFolderName name, string? mailbox) =>
        string.IsNullOrWhiteSpace(mailbox) ? new FolderId(name) : new FolderId(name, new Mailbox(mailbox.Trim()));

    /// <summary>
    /// Resolves a folder given as well-known name (inbox, sentitems, drafts, deleteditems, junkemail, archive, ...),
    /// as path below the mailbox root ("Inbox/Projects/Alpha"), as display name, or as EWS folder id.
    /// </summary>
    public async Task<FolderId> ResolveFolderAsync(string? folder, string? mailbox, WellKnownFolderName defaultFolder)
    {
        if (string.IsNullOrWhiteSpace(folder)) return WellKnown(defaultFolder, mailbox);
        folder = folder.Trim();

        if (TryWellKnown(folder, out var wellKnown)) return WellKnown(wellKnown, mailbox);

        if (folder.Contains('/'))
        {
            var segments = folder.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var current = TryWellKnown(segments[0], out var first) ? WellKnown(first, mailbox) : null;
            var start = current is null ? 0 : 1;
            current ??= WellKnown(WellKnownFolderName.MsgFolderRoot, mailbox);
            for (var i = start; i < segments.Length; i++)
            {
                var found = await Service.FindFolders(current,
                    new SearchFilter.IsEqualTo(FolderSchema.DisplayName, segments[i]),
                    new FolderView(2) { PropertySet = new PropertySet(BasePropertySet.IdOnly) });
                current = found.Folders.FirstOrDefault()?.Id
                    ?? throw new McpException($"Folder '{segments[i]}' not found in path '{folder}'.");
            }
            return current;
        }

        var byName = await Service.FindFolders(WellKnown(WellKnownFolderName.MsgFolderRoot, mailbox),
            new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folder),
            new FolderView(5) { Traversal = FolderTraversal.Deep, PropertySet = new PropertySet(BasePropertySet.IdOnly) });
        if (byName.Folders.Count == 1) return byName.Folders[0].Id;
        if (byName.Folders.Count > 1)
        {
            throw new McpException($"Folder name '{folder}' is ambiguous; use a path like 'Inbox/{folder}' or the folder id from list_folders.");
        }

        // Treat as an EWS folder id.
        return new FolderId(folder);
    }

    private static bool TryWellKnown(string value, out WellKnownFolderName name)
    {
        var normalized = value.Replace(" ", "").Replace("_", "").ToLowerInvariant();
        normalized = normalized switch
        {
            "sent" or "sentmail" => "sentitems",
            "deleted" or "trash" => "deleteditems",
            "junk" or "spam" => "junkemail",
            "draft" => "drafts",
            "archive" => "archivemsgfolderroot",
            "root" => "msgfolderroot",
            _ => normalized,
        };
        return Enum.TryParse(normalized, ignoreCase: true, out name) && Enum.IsDefined(name);
    }

    // ---- Date handling -------------------------------------------------------------------------

    /// <summary>
    /// Parses an ISO 8601 date/time. Values with offset ("Z", "+02:00") are absolute; values without offset are
    /// interpreted in the configured time zone.
    /// </summary>
    public DateTime ParseDateTime(string value, string parameterName) => ParseDateTimeValue(value, parameterName);

    internal static DateTime ParseDateTimeValue(string value, string parameterName)
    {
        value = value.Trim();
        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dto)
            && (value.EndsWith('Z') || value.EndsWith('z') || HasOffset(value)))
        {
            return dto.UtcDateTime;
        }
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var local))
        {
            return DateTime.SpecifyKind(local, DateTimeKind.Unspecified);
        }
        throw new McpException($"Invalid date/time for '{parameterName}': '{value}'. Use ISO 8601, e.g. 2026-11-03T14:30.");
    }

    private static bool HasOffset(string value)
    {
        var t = value.IndexOf('T');
        if (t < 0) return false;
        var time = value[(t + 1)..];
        return time.Contains('+') || time.Contains('-');
    }

    /// <summary>Current date/time in the configured time zone.</summary>
    public DateTime Now => TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZone);

    /// <summary>Formats a DateTime returned by EWS (expressed in the service time zone) as ISO 8601 with offset.</summary>
    public string Format(DateTime value)
    {
        var dt = value.Kind == DateTimeKind.Utc
            ? TimeZoneInfo.ConvertTimeFromUtc(value, TimeZone)
            : DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        return new DateTimeOffset(dt, TimeZone.GetUtcOffset(dt)).ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture);
    }

    public string? TryFormat(Func<DateTime> getter)
    {
        try
        {
            return Format(getter());
        }
        catch (ServiceObjectPropertyException)
        {
            return null;
        }
    }

    public void EnsureSendAllowed(string what)
    {
        if (!Options.AllowSend)
        {
            throw new McpException($"Sending is disabled on this server ({what}). The item was not sent; the user can review and send it in Outlook.");
        }
    }
}
