using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using ExchangeEwsMcp.Configuration;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

public sealed record IssuedTokens(string AccessToken, int ExpiresIn, string RefreshToken);

/// <summary>Thrown when a grant can no longer be used (revoked, upstream refresh failed). The user must sign in again.</summary>
public sealed class GrantExpiredException(string message) : Exception(message);

/// <summary>
/// Issues the server's own access / refresh tokens to MCP clients and provides valid upstream (EWS) access tokens for a grant.
/// </summary>
public sealed class TokenService(
    OAuthStore store,
    UpstreamIdentityProvider upstream,
    IDataProtectionProvider dataProtection,
    IOptions<OAuthServerOptions> options,
    IOptions<ServerOptions> server,
    TimeProvider time,
    ILogger<TokenService> logger)
{
    private static readonly TimeSpan UpstreamRefreshSkew = TimeSpan.FromMinutes(3);

    private readonly IDataProtector _accessTokenProtector = dataProtection.CreateProtector("ExchangeEwsMcp.AccessToken.v1");
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _refreshLocks = new();

    /// <summary>Issues a new access token and rotates the refresh token of the grant.</summary>
    public IssuedTokens Issue(Grant grant)
    {
        var lifetime = options.Value.AccessTokenLifetime;
        var payload = new AccessTokenPayload(grant.Id, grant.ClientId, server.Value.McpResource, time.GetUtcNow().Add(lifetime).ToUnixTimeSeconds());
        var accessToken = _accessTokenProtector.Protect(JsonSerializer.Serialize(payload));

        var refreshToken = Crypto.RandomToken();
        grant.RefreshTokenHash = Crypto.Sha256(refreshToken);
        grant.RefreshTokenExpiresAt = time.GetUtcNow().Add(options.Value.RefreshTokenLifetime);
        store.SaveGrant(grant);

        return new IssuedTokens(accessToken, (int)lifetime.TotalSeconds, refreshToken);
    }

    /// <summary>Validates an access token presented by an MCP client and returns its grant.</summary>
    public Grant? ValidateAccessToken(string token)
    {
        AccessTokenPayload? payload;
        try
        {
            payload = JsonSerializer.Deserialize<AccessTokenPayload>(_accessTokenProtector.Unprotect(token));
        }
        catch (Exception e) when (e is CryptographicException or JsonException or FormatException)
        {
            return null;
        }

        if (payload is null
            || payload.Exp < time.GetUtcNow().ToUnixTimeSeconds()
            || !string.Equals(payload.Aud, server.Value.McpResource, StringComparison.Ordinal))
        {
            return null;
        }

        var grant = store.FindGrant(payload.Gid);
        return grant is { Revoked: false } && grant.ClientId == payload.Cid ? grant : null;
    }

    /// <summary>Returns a valid upstream access token for EWS, refreshing it if required.</summary>
    public async Task<string> GetUpstreamAccessTokenAsync(string grantId, CancellationToken ct)
    {
        var grant = store.FindGrant(grantId) ?? throw new GrantExpiredException("Unknown grant.");
        if (IsUsable(grant)) return grant.UpstreamAccessToken!;

        var gate = _refreshLocks.GetOrAdd(grantId, _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(ct);
        try
        {
            grant = store.FindGrant(grantId) ?? throw new GrantExpiredException("Unknown grant.");
            if (IsUsable(grant)) return grant.UpstreamAccessToken!;
            if (grant.Revoked) throw new GrantExpiredException("The authorization was revoked. Please sign in again.");
            if (string.IsNullOrEmpty(grant.UpstreamRefreshToken))
            {
                throw new GrantExpiredException("The Exchange session expired and cannot be renewed. Please sign in again.");
            }

            UpstreamTokens tokens;
            try
            {
                tokens = await upstream.RefreshAsync(grant.UpstreamRefreshToken, ct);
            }
            catch (UpstreamAuthException e) when (e.Error is "invalid_grant" or "interaction_required" or "invalid_client" or "unauthorized_client")
            {
                logger.LogInformation("Upstream refresh for grant {GrantId} failed ({Error}); revoking grant.", grantId, e.Error);
                grant.Revoked = true;
                store.SaveGrant(grant);
                throw new GrantExpiredException("The Exchange session expired. Please sign in again.");
            }

            grant.UpstreamAccessToken = tokens.AccessToken;
            grant.UpstreamAccessTokenExpiresAt = tokens.ExpiresAt;
            if (!string.IsNullOrEmpty(tokens.RefreshToken))
            {
                grant.UpstreamRefreshToken = tokens.RefreshToken;
            }
            store.SaveGrant(grant);
            return tokens.AccessToken;
        }
        finally
        {
            gate.Release();
        }
    }

    private bool IsUsable(Grant grant) =>
        !grant.Revoked
        && !string.IsNullOrEmpty(grant.UpstreamAccessToken)
        && grant.UpstreamAccessTokenExpiresAt - UpstreamRefreshSkew > time.GetUtcNow();

    private sealed record AccessTokenPayload(string Gid, string Cid, string Aud, long Exp);
}
