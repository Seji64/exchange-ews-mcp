using System.Security.Claims;
using System.Text.Encodings.Web;
using ExchangeEwsMcp.Configuration;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

/// <summary>Authenticates MCP requests with bearer tokens issued by the built-in authorization server.</summary>
public sealed class AccessTokenAuthenticationHandler(
    IOptionsMonitor<AuthenticationSchemeOptions> options,
    ILoggerFactory logger,
    UrlEncoder encoder,
    TokenService tokens,
    IOptions<OAuthServerOptions> oauth)
    : AuthenticationHandler<AuthenticationSchemeOptions>(options, logger, encoder)
{
    public const string SchemeName = "ExchangeEwsMcpBearer";
    public const string GrantIdClaim = "grant_id";
    public const string PassthroughTokenClaim = "ews_access_token";

    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        string? header = Request.Headers.Authorization;
        if (string.IsNullOrEmpty(header) || !header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            return Task.FromResult(AuthenticateResult.NoResult());
        }

        var token = header["Bearer ".Length..].Trim();
        if (oauth.Value.Mode == AuthMode.Passthrough)
        {
            return Task.FromResult(AuthenticatePassthrough(token));
        }

        var grant = tokens.ValidateAccessToken(token);
        if (grant is null)
        {
            return Task.FromResult(AuthenticateResult.Fail("Invalid or expired access token."));
        }

        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, grant.Subject),
            new(GrantIdClaim, grant.Id),
            new("client_id", grant.ClientId),
        };
        if (grant.UserPrincipalName is not null) claims.Add(new Claim(ClaimTypes.Upn, grant.UserPrincipalName));
        if (grant.Email is not null) claims.Add(new Claim(ClaimTypes.Email, grant.Email));
        if (grant.DisplayName is not null) claims.Add(new Claim(ClaimTypes.Name, grant.DisplayName));

        return Task.FromResult(Success(claims));
    }

    /// <summary>
    /// Exchange validates the token's signature and audience on every EWS call; the claims are only read here for
    /// expiry (so the client refreshes in time) and display. Caches are keyed by the token hash, not by claims.
    /// </summary>
    private AuthenticateResult AuthenticatePassthrough(string token)
    {
        var payload = UpstreamIdentityProvider.ReadJwtPayload(token);
        if (payload is null
            || !payload.TryGetValue("exp", out var exp) || !exp.TryGetInt64(out var expSeconds)
            || expSeconds < TimeProvider.GetUtcNow().ToUnixTimeSeconds())
        {
            return AuthenticateResult.Fail("Invalid or expired access token.");
        }

        UpstreamUser user;
        try
        {
            user = UpstreamIdentityProvider.ReadUser(new UpstreamTokens(token, default, null, null));
        }
        catch (UpstreamAuthException)
        {
            return AuthenticateResult.Fail("Access token contains no user.");
        }

        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, user.Subject),
            new(GrantIdClaim, "token:" + Crypto.Sha256(token)),
            new(PassthroughTokenClaim, token),
        };
        if (user.UserPrincipalName is not null) claims.Add(new Claim(ClaimTypes.Upn, user.UserPrincipalName));
        if (user.Email is not null) claims.Add(new Claim(ClaimTypes.Email, user.Email));
        if (user.DisplayName is not null) claims.Add(new Claim(ClaimTypes.Name, user.DisplayName));
        return Success(claims);
    }

    private static AuthenticateResult Success(List<Claim> claims) =>
        AuthenticateResult.Success(new AuthenticationTicket(
            new ClaimsPrincipal(new ClaimsIdentity(claims, SchemeName, ClaimTypes.Name, ClaimTypes.Role)), SchemeName));
}
