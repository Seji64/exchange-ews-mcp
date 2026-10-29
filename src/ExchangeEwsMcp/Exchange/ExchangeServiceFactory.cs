using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
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
    private readonly X509Certificate2Collection? _trustedCas;

    public ExchangeServiceFactory(IOptions<ExchangeOptions> options, TokenService tokens, ILogger<ExchangeServiceFactory> logger)
    {
        _options = options.Value;
        _tokens = tokens;
        _logger = logger;
        TimeZone = ResolveTimeZone(_options.TimeZone);
        if (!string.IsNullOrEmpty(_options.TrustedCaCertificatePath))
        {
            _trustedCas = [];
            _trustedCas.ImportFromPemFile(_options.TrustedCaCertificatePath);
        }
    }

    /// <summary>Time zone used for the EWS requests and for rendering dates.</summary>
    public TimeZoneInfo TimeZone { get; }

    public ExchangeService GetService(string grantId) =>
        _cache.GetOrCreate(grantId, entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(30);
            entry.RegisterPostEvictionCallback(static (_, value, _, _) => (value as IDisposable)?.Dispose());
            return Create(grantId);
        })!;

    private ExchangeService Create(string grantId)
    {
        var service = new ExchangeService(ExchangeVersion.Exchange2016, TimeZone)
        {
            Url = new Uri(_options.EwsUrl),
            Credentials = new GrantCredentials(_tokens, grantId),
            UseDefaultCredentials = false,
            AcceptGzipEncoding = true,
            UserAgent = "exchange-ews-mcp",
            Timeout = 100_000,
        };
        if (_options.IgnoreCertificateErrors)
        {
            service.ServerCertificateValidationCallback = (_, _, _, _) => true;
        }
        else if (_trustedCas is not null)
        {
            service.ServerCertificateValidationCallback = ValidateWithTrustedCas;
        }
        if (_options.TraceEws)
        {
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            service.TraceListener = new LoggerTraceListener(_logger);
        }
        return service;
    }

    private bool ValidateWithTrustedCas(HttpRequestMessage _, X509Certificate2? certificate, X509Chain? __, SslPolicyErrors errors)
    {
        if (errors == SslPolicyErrors.None) return true;
        if (certificate is null || errors.HasFlag(SslPolicyErrors.RemoteCertificateNameMismatch)) return false;

        using var chain = new X509Chain();
        chain.ChainPolicy.TrustMode = X509ChainTrustMode.CustomRootTrust;
        chain.ChainPolicy.CustomTrustStore.AddRange(_trustedCas!);
        chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
        return chain.Build(certificate);
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
