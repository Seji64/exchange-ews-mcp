using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExchangeEwsMcp.Configuration;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

public sealed record UpstreamTokens(string AccessToken, DateTimeOffset ExpiresAt, string? RefreshToken, string? IdToken);

public sealed record UpstreamUser(string Subject, string? UserPrincipalName, string? Email, string? DisplayName, string? Nonce);

public sealed class UpstreamAuthException(string error, string? description) : Exception($"{error}: {description}")
{
    public string Error { get; } = error;
    public string? Description { get; } = description;
}

/// <summary>
/// Minimal OpenID Connect / OAuth 2.0 client for the upstream identity provider (Entra ID or AD FS).
/// Uses the authorization code flow with PKCE and a confidential client secret.
/// </summary>
public sealed class UpstreamIdentityProvider(
    IHttpClientFactory httpClientFactory,
    IOptions<IdentityProviderOptions> options,
    IOptions<ServerOptions> server,
    TimeProvider time,
    ILogger<UpstreamIdentityProvider> logger)
{
    public const string HttpClientName = "upstream-idp";

    private readonly IdentityProviderOptions _options = options.Value;
    private readonly SemaphoreSlim _discoveryLock = new(1, 1);
    private DiscoveryDocument? _discovery;
    private DateTimeOffset _discoveryLoadedAt;

    public string CallbackUri => server.Value.Issuer + "/oauth/callback";

    public async Task<string> BuildAuthorizeUrlAsync(string state, string nonce, string codeVerifier, CancellationToken ct)
    {
        var discovery = await GetDiscoveryAsync(ct);
        var query = new Dictionary<string, string?>
        {
            ["client_id"] = _options.ClientId,
            ["response_type"] = "code",
            ["response_mode"] = "query",
            ["redirect_uri"] = CallbackUri,
            ["scope"] = string.Join(' ', _options.Scopes),
            ["state"] = state,
            ["nonce"] = nonce,
            ["code_challenge"] = Crypto.PkceChallenge(codeVerifier),
            ["code_challenge_method"] = "S256",
        };
        if (!string.IsNullOrEmpty(_options.Resource))
        {
            query["resource"] = _options.Resource;
        }
        return QueryHelpers.AddQueryString(discovery.AuthorizationEndpoint, query);
    }

    public Task<UpstreamTokens> RedeemCodeAsync(string code, string codeVerifier, CancellationToken ct) =>
        RequestTokenAsync(new Dictionary<string, string>
        {
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = CallbackUri,
            ["code_verifier"] = codeVerifier,
        }, ct);

    public Task<UpstreamTokens> RefreshAsync(string refreshToken, CancellationToken ct) =>
        RequestTokenAsync(new Dictionary<string, string>
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = refreshToken,
        }, ct);

    private async Task<UpstreamTokens> RequestTokenAsync(Dictionary<string, string> form, CancellationToken ct)
    {
        var discovery = await GetDiscoveryAsync(ct);
        form["client_id"] = _options.ClientId;
        form["client_secret"] = _options.ClientSecret;
        form["scope"] = string.Join(' ', _options.Scopes);
        if (!string.IsNullOrEmpty(_options.Resource))
        {
            form["resource"] = _options.Resource;
        }

        var http = httpClientFactory.CreateClient(HttpClientName);
        using var response = await http.PostAsync(discovery.TokenEndpoint, new FormUrlEncodedContent(form), ct);
        var body = await response.Content.ReadAsStringAsync(ct);
        TokenResponse? token = null;
        try
        {
            token = JsonSerializer.Deserialize<TokenResponse>(body);
        }
        catch (JsonException)
        {
        }

        if (!response.IsSuccessStatusCode || token?.AccessToken is null)
        {
            logger.LogWarning("Upstream token request ({GrantType}) failed with {Status}: {Error} {Description}",
                form["grant_type"], (int)response.StatusCode, token?.Error, token?.ErrorDescription);
            throw new UpstreamAuthException(token?.Error ?? "token_request_failed", token?.ErrorDescription ?? $"HTTP {(int)response.StatusCode}");
        }

        var expiresIn = token.ExpiresIn ?? 3600;
        return new UpstreamTokens(token.AccessToken, time.GetUtcNow().AddSeconds(expiresIn), token.RefreshToken, token.IdToken);
    }

    /// <summary>
    /// Extracts the user identity from the ID token (falls back to the access token).
    /// The tokens were received directly from the token endpoint over TLS, so the signature is not re-validated
    /// (OpenID Connect Core 3.1.3.7).
    /// </summary>
    public static UpstreamUser ReadUser(UpstreamTokens tokens)
    {
        var claims = ReadJwtPayload(tokens.IdToken) ?? ReadJwtPayload(tokens.AccessToken)
            ?? throw new UpstreamAuthException("invalid_token", "Upstream identity provider returned no readable ID token.");

        string? Claim(params string[] names)
        {
            foreach (var name in names)
            {
                if (claims.TryGetValue(name, out var v) && v.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v.GetString()))
                {
                    return v.GetString();
                }
            }
            return null;
        }

        var subject = Claim("oid", "sub", "upn", "unique_name")
            ?? throw new UpstreamAuthException("invalid_token", "ID token contains no subject.");
        var upn = Claim("upn", "preferred_username", "unique_name");
        var email = Claim("email", "preferred_username", "upn");
        return new UpstreamUser(subject, upn, email, Claim("name", "given_name"), Claim("nonce"));
    }

    internal static Dictionary<string, JsonElement>? ReadJwtPayload(string? jwt)
    {
        if (string.IsNullOrEmpty(jwt)) return null;
        var parts = jwt.Split('.');
        if (parts.Length < 2) return null;
        try
        {
            var json = Encoding.UTF8.GetString(Crypto.FromBase64Url(parts[1]));
            return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);
        }
        catch (Exception e) when (e is FormatException or JsonException)
        {
            return null;
        }
    }

    private async Task<DiscoveryDocument> GetDiscoveryAsync(CancellationToken ct)
    {
        if (_discovery is not null && time.GetUtcNow() - _discoveryLoadedAt < TimeSpan.FromHours(12))
        {
            return _discovery;
        }

        await _discoveryLock.WaitAsync(ct);
        try
        {
            if (_discovery is not null && time.GetUtcNow() - _discoveryLoadedAt < TimeSpan.FromHours(12))
            {
                return _discovery;
            }
            var url = _options.Authority.TrimEnd('/') + "/.well-known/openid-configuration";
            var http = httpClientFactory.CreateClient(HttpClientName);
            var doc = await http.GetFromJsonAsync<DiscoveryDocument>(url, ct)
                ?? throw new InvalidOperationException($"Empty discovery document at {url}");
            if (string.IsNullOrEmpty(doc.AuthorizationEndpoint) || string.IsNullOrEmpty(doc.TokenEndpoint))
            {
                throw new InvalidOperationException($"Discovery document at {url} lacks authorization/token endpoints.");
            }
            _discovery = doc;
            _discoveryLoadedAt = time.GetUtcNow();
            return doc;
        }
        finally
        {
            _discoveryLock.Release();
        }
    }

    private sealed class DiscoveryDocument
    {
        [JsonPropertyName("authorization_endpoint")]
        public string AuthorizationEndpoint { get; set; } = "";

        [JsonPropertyName("token_endpoint")]
        public string TokenEndpoint { get; set; } = "";
    }

    private sealed class TokenResponse
    {
        [JsonPropertyName("access_token")]
        public string? AccessToken { get; set; }

        [JsonPropertyName("refresh_token")]
        public string? RefreshToken { get; set; }

        [JsonPropertyName("id_token")]
        public string? IdToken { get; set; }

        [JsonPropertyName("expires_in")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? ExpiresIn { get; set; }

        [JsonPropertyName("error")]
        public string? Error { get; set; }

        [JsonPropertyName("error_description")]
        public string? ErrorDescription { get; set; }
    }
}
