using System.Security.Claims;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

/// <summary>Authenticates MCP requests with bearer tokens issued by the built-in authorization server.</summary>
public sealed class AccessTokenAuthenticationHandler(
    IOptionsMonitor<AuthenticationSchemeOptions> options,
    ILoggerFactory logger,
    UrlEncoder encoder,
    TokenService tokens)
    : AuthenticationHandler<AuthenticationSchemeOptions>(options, logger, encoder)
{
    public const string SchemeName = "ExchangeEwsMcpBearer";
    public const string GrantIdClaim = "grant_id";

    protected override Task<AuthenticateResult> HandleAuthenticateAsync()
    {
        string? header = Request.Headers.Authorization;
        if (string.IsNullOrEmpty(header) || !header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            return Task.FromResult(AuthenticateResult.NoResult());
        }

        var token = header["Bearer ".Length..].Trim();
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

        var identity = new ClaimsIdentity(claims, SchemeName, ClaimTypes.Name, ClaimTypes.Role);
        return Task.FromResult(AuthenticateResult.Success(new AuthenticationTicket(new ClaimsPrincipal(identity), SchemeName)));
    }
}
