using ExchangeEwsMcp.Auth;
using ExchangeEwsMcp.Configuration;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Exchange;

/// <summary>
/// Creates and caches one <see cref="ExchangeService"/> per grant. The service fetches a fresh upstream
/// access token for every EWS request, so tokens are refreshed transparently.
/// </summary>
public sealed class ExchangeServiceFactory : IDisposable
{
    private readonly ExchangeOptions _options;
    private readonly TokenService _tokens;
    private readonly ILogger<ExchangeServiceFactory> _logger;
    private readonly MemoryCache _cache = new(new MemoryCacheOptions());

    public ExchangeServiceFactory(IOptions<ExchangeOptions> options, TokenService tokens, ILogger<ExchangeServiceFactory> logger)
    {
        _options = options.Value;
        _tokens = tokens;
        _logger = logger;
        TimeZone = ResolveTimeZone(_options.TimeZone);
    }

    /// <summary>Time zone used for the EWS requests and for rendering dates.</summary>
    public TimeZoneInfo TimeZone { get; }

    /// <param name="grantId">Cache key: the grant id, or the token hash in passthrough mode.</param>
    /// <param name="passthroughToken">In passthrough mode, the EWS access token sent by the client.</param>
    /// <param name="anchorMailbox">Mailbox the requests are routed to (X-AnchorMailbox); null for the user's own.</param>
    public ExchangeService GetService(string grantId, string? passthroughToken = null, string? anchorMailbox = null) =>
        _cache.GetOrCreate(grantId + "|" + anchorMailbox?.ToLowerInvariant(), entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(30);
            entry.RegisterPostEvictionCallback(static (_, value, _, _) => (value as IDisposable)?.Dispose());
            var service = Create(passthroughToken is null ? new GrantCredentials(_tokens, grantId) : new OAuthCredentials(passthroughToken));
            if (anchorMailbox is not null)
            {
                service.HttpHeaders["X-AnchorMailbox"] = anchorMailbox;
            }
            return service;
        })!;

    private ExchangeService Create(ExchangeCredentials credentials)
    {
        var service = new ExchangeService(ExchangeVersion.Exchange2016, TimeZone)
        {
            Url = new Uri(_options.EwsUrl),
            Credentials = credentials,
            UseDefaultCredentials = false,
            AcceptGzipEncoding = true,
            UserAgent = "exchange-ews-mcp",
            Timeout = 100_000,
        };
        if (_options.TraceEws)
        {
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            service.TraceListener = new LoggerTraceListener(_logger);
        }
        return service;
    }

    /// <summary>
    /// Resolves the configured time zone. Exchange expects Windows time zone ids, so IANA ids are mapped
    /// to their Windows equivalent where possible.
    /// </summary>
    internal static TimeZoneInfo ResolveTimeZone(string id)
    {
        if (string.IsNullOrWhiteSpace(id)) return TimeZoneInfo.Utc;
        if (TimeZoneInfo.TryConvertIanaIdToWindowsId(id, out var windowsId))
        {
            id = windowsId;
        }
        return TimeZoneInfo.FindSystemTimeZoneById(id);
    }

    public void Dispose() => _cache.Dispose();

    private sealed class GrantCredentials(TokenService tokens, string grantId) : OAuthAccessTokenCredentials
    {
        public override Task<string> AcquireAccessToken() => tokens.GetUpstreamAccessTokenAsync(grantId, CancellationToken.None);
    }

    private sealed class LoggerTraceListener(ILogger logger) : ITraceListener
    {
        public void Trace(string traceType, string traceMessage) => logger.LogDebug("EWS {TraceType}: {Message}", traceType, traceMessage);
    }
}
