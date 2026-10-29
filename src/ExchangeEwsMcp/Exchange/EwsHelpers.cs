using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol;

namespace ExchangeEwsMcp.Exchange;

internal static class EwsHelpers
{
    public static PersonInfo? ToPerson(EmailAddress? address) =>
        address is null ? null : new PersonInfo(NullIfEmpty(address.Name), NullIfEmpty(address.Address));

    public static IReadOnlyList<PersonInfo> ToPeople(IEnumerable<EmailAddress>? addresses) =>
        addresses?.Select(a => ToPerson(a)!).ToList() ?? [];

    public static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    /// <summary>Reads an optional property without throwing if it was not loaded / is not set.</summary>
    public static T? Try<T>(Func<T> getter)
    {
        try
        {
            return getter();
        }
        catch (Exception e) when (e is ServiceObjectPropertyException or ServiceVersionException or KeyNotFoundException or NullReferenceException)
        {
            return default;
        }
    }

    public static (string Text, bool Truncated) Truncate(string? text, int max)
    {
        text ??= "";
        return text.Length <= max ? (text, false) : (text[..max], true);
    }

    /// <summary>Builds a message body. Plain text is converted to simple HTML so it renders well in Outlook.</summary>
    public static MessageBody CreateBody(string? body, string? format)
    {
        body ??= "";
        if (string.Equals(format, "html", StringComparison.OrdinalIgnoreCase))
        {
            return new MessageBody(BodyType.HTML, body);
        }
        var html = new StringBuilder("<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\">");
        foreach (var line in body.Replace("\r\n", "\n").Split('\n'))
        {
            html.Append(line.Length == 0 ? "<br>" : $"<div>{WebUtility.HtmlEncode(line)}</div>");
        }
        html.Append("</div>");
        return new MessageBody(BodyType.HTML, html.ToString());
    }

    /// <summary>
    /// Resolves recipients given as e-mail address or (partial) name. Names must resolve to exactly one entry,
    /// otherwise the candidates are reported back so the model can ask the user.
    /// </summary>
    public static async Task<List<EmailAddress>> ResolveRecipientsAsync(ExchangeService service, IEnumerable<string>? recipients)
    {
        var result = new List<EmailAddress>();
        if (recipients is null) return result;
        foreach (var raw in recipients)
        {
            var value = raw?.Trim();
            if (string.IsNullOrEmpty(value)) continue;

            var angle = value.IndexOf('<');
            if (angle >= 0 && value.EndsWith('>'))
            {
                result.Add(new EmailAddress(value[..angle].Trim().Trim('"'), value[(angle + 1)..^1].Trim()));
                continue;
            }
            if (value.Contains('@'))
            {
                result.Add(new EmailAddress(value));
                continue;
            }

            var matches = await service.ResolveName(value, ResolveNameSearchLocation.ContactsThenDirectory, false);
            var mailboxes = matches.Where(m => !string.IsNullOrEmpty(m.Mailbox.Address)).ToList();
            if (mailboxes.Count == 1)
            {
                result.Add(new EmailAddress(mailboxes[0].Mailbox.Name, mailboxes[0].Mailbox.Address));
                continue;
            }
            if (mailboxes.Count == 0)
            {
                throw new McpException($"Recipient '{value}' could not be resolved. Use an e-mail address or call resolve_names.");
            }
            var candidates = string.Join("; ", mailboxes.Take(10).Select(m => $"{m.Mailbox.Name} <{m.Mailbox.Address}>"));
            throw new McpException($"Recipient '{value}' is ambiguous. Ask the user which one is meant: {candidates}");
        }
        return result;
    }

    public static string DescribeEwsError(Exception e) => e switch
    {
        ServiceResponseException sre => sre.ErrorCode switch
        {
            ServiceError.ErrorItemNotFound => "The item was not found (it may have been moved or deleted, or the id is invalid).",
            ServiceError.ErrorFolderNotFound => "The folder was not found.",
            ServiceError.ErrorAccessDenied =>
                "Access denied. The user lacks permission for this mailbox or item.",
            ServiceError.ErrorNonExistentMailbox => "The mailbox does not exist.",
            ServiceError.ErrorInvalidIdMalformed or ServiceError.ErrorInvalidIdNotAnItemAttachmentId => "The id is malformed.",
            _ => $"Exchange returned {sre.ErrorCode}: {sre.Message}",
        },
        ServiceRequestException when HttpStatus(e) == HttpStatusCode.Unauthorized =>
            "Exchange rejected the access token (HTTP 401). Check the Hybrid Modern Authentication / AD FS configuration " +
            "(EWS URL registered as service principal name / AD FS Web API identifier, EWS.AccessAsUser.All scope).",
        ServiceRequestException when HttpStatus(e) is { } status =>
            $"Exchange returned HTTP {(int)status} {status}.",
        ServiceRequestException sre => $"The request to Exchange failed: {sre.Message}",
        ServiceLocalException sle => sle.Message,
        _ => e.Message,
    };

    private static HttpStatusCode? HttpStatus(Exception e)
    {
        for (var inner = e.InnerException; inner is not null; inner = inner.InnerException)
        {
            // EwsHttpClientException is internal to the EWS library; read its Response property via reflection.
            if (inner.GetType().GetProperty("Response")?.GetValue(inner) is HttpResponseMessage response) return response.StatusCode;
            if (inner is HttpRequestException { StatusCode: { } status }) return status;
        }
        return null;
    }
}
